VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCDToc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" ( _
      lpEventAttributes As Any, _
      ByVal bManualReset As Long, _
      ByVal bInitialState As Long, _
      lpName As Any) As Long
Private Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

'*******************************************************************
'** TOC format
'*******************************************************************
Private Type TOC_TRACK
    Rsvd1 As Byte
    ADR As Byte
    Track As Byte
    Rsvd2 As Byte
    Addr(3) As Byte
End Type

Private Type TOC
    TocLen(1) As Byte
    FirstTrack As Byte
    LastTrack As Byte
    tocTrack(99) As TOC_TRACK
End Type

Public Enum ECDTrackTypes
   ECDAudio = 16
   ECDData = 20
End Enum

Private m_iTracks As Long
Private m_tTrackInfo As TOC
Private m_iCDDBID As Long
Private m_bNoCD As Boolean

Private m_lPower2(0 To 31) As Long

Public Property Get NoCD() As Boolean
   NoCD = m_bNoCD
End Property

Public Property Get TrackCount() As Long
   If (m_tTrackInfo.FirstTrack = 0) And (m_tTrackInfo.LastTrack = 0) Then
      TrackCount = 0
   Else
      TrackCount = m_tTrackInfo.LastTrack - m_tTrackInfo.FirstTrack + 1
   End If
End Property

Public Property Get CDDBId() As Long
   CDDBId = m_iCDDBID
End Property

Public Property Get CDDBQuery() As String
Dim sQuery As String
   sQuery = "cddb query"
   sQuery = sQuery & " " & Hex$(m_iCDDBID)
   sQuery = sQuery & " " & TrackCount
Dim i As Long
   For i = 1 To TrackCount()
      sQuery = sQuery & " " & TrackOffset(i)
   Next i
   sQuery = sQuery & " " & TotalLengthSeconds
   CDDBQuery = LCase(sQuery)
End Property

Public Property Get TrackOffset(ByVal index As Long) As Long
   TrackOffset = ((TrackLengthMinutes(index) * 60) + TrackLengthSeconds(index)) * 75 + TrackLengthFrames(index)
End Property

Public Property Get TrackLengthMinutes(ByVal index As Long) As Long
   TrackLengthMinutes = m_tTrackInfo.tocTrack(index - 1).Addr(1)
End Property

Public Property Get TrackLengthSeconds(ByVal index As Long) As Long
   TrackLengthSeconds = m_tTrackInfo.tocTrack(index - 1).Addr(2)
End Property

Public Property Get TrackLengthFrames(ByVal index As Long) As Long
   TrackLengthFrames = m_tTrackInfo.tocTrack(index - 1).Addr(3)
End Property

Public Property Get TrackType(ByVal index As Long) As ECDTrackTypes
   TrackType = m_tTrackInfo.tocTrack(index - 1).ADR
End Property

Public Property Get LeadOutOffset() As Long
Dim leadOutIndex As Long
   leadOutIndex = m_tTrackInfo.LastTrack - m_tTrackInfo.FirstTrack + 1
   LeadOutOffset = TrackOffset(leadOutIndex + 1)
End Property
   
Public Property Get LeadOutMinutes() As Long
Dim leadOutIndex As Long
   leadOutIndex = m_tTrackInfo.LastTrack - m_tTrackInfo.FirstTrack + 1
   LeadOutMinutes = TrackLengthMinutes(leadOutIndex + 1)
End Property

Public Property Get LeadOutSeconds() As Long
Dim leadOutIndex As Long
   leadOutIndex = m_tTrackInfo.LastTrack - m_tTrackInfo.FirstTrack + 1
   LeadOutSeconds = TrackLengthSeconds(leadOutIndex + 1)
End Property

Public Property Get LeadOutFrames() As Long
Dim leadOutIndex As Long
   leadOutIndex = m_tTrackInfo.LastTrack - m_tTrackInfo.FirstTrack + 1
   LeadOutFrames = TrackLengthFrames(leadOutIndex + 1)
End Property

Public Property Get TotalLengthSeconds() As Long
   TotalLengthSeconds = LeadOutOffset \ 75
End Property

Friend Sub fInit( _
      ByVal haId As Long, _
      ByVal id As Long, _
      ByVal lun As Long _
   )
   ' Clear any existing info:
   m_tTrackInfo.FirstTrack = 0
   m_tTrackInfo.LastTrack = 0
   m_iCDDBID = 0
   m_bNoCD = True


Dim tToc As TOC
Dim ExecIO As SRB_ExecuteIO
Dim lR As Long
Dim hEvent As Long

   hEvent = CreateEvent(ByVal 0&, 1, 0, ByVal 0&)

   ' build the toc:
   ExecIO.SRB_Cmd = SC_EXEC_SCSI_CMD
   
   ExecIO.SRB_HaID = haId
   ExecIO.SRB_Target = id
   ExecIO.SRB_Lun = lun
   
   ExecIO.SRB_Flags = SRB_DIR_IN
   
   ExecIO.SRB_BufLen = &H324
   ExecIO.SRB_BufPointer = VarPtr(tToc)

   ExecIO.SRB_SenseLen = &HE
   ExecIO.SRB_CDBLen = &HA
   ExecIO.SRB_CDBByte(0) = &H43    'read TOC command
   ExecIO.SRB_CDBByte(1) = &H2     'MSF mode
   ExecIO.SRB_CDBByte(7) = &H3     'high-order byte of buffer len
   ExecIO.SRB_CDBByte(8) = &H24    'low-order byte of buffer len
   
   ExecIO.SRB_PostProc = hEvent

   ResetEvent hEvent
   lR = SendASPI32ExecIOEx(ExecIO)
   'While ExecIO.SRB_Status = SS_PENDING
   '     DoEvents
   'Wend
   If (lR = SS_PENDING) Then
      WaitForSingleObject hEvent, 10000
   End If
   CloseHandle hEvent
   
   If Not (ExecIO.SRB_Status = SS_COMP) Then
      Debug.Print "error IO"
   Else
      m_bNoCD = False
      LSet m_tTrackInfo = tToc
   
      getCDDBID
   End If

End Sub

Private Sub getCDDBID()
Dim i As Long
Dim numTracks As Long
Dim tocTrack As TOC_TRACK
Dim n As Long
Dim t As Long
Dim Res As Long

   numTracks = m_tTrackInfo.LastTrack - m_tTrackInfo.FirstTrack + 1
   For i = 0 To numTracks - 1
      tocTrack = m_tTrackInfo.tocTrack(i)
      n = UnsignedAdd(n, CDDBSum(60 * tocTrack.Addr(1) + tocTrack.Addr(2)))
   Next i
   
   tocTrack = m_tTrackInfo.tocTrack(numTracks)
   t = 60 * tocTrack.Addr(1) + tocTrack.Addr(2)
   tocTrack = m_tTrackInfo.tocTrack(0)
   t = t - (60 * tocTrack.Addr(1) + tocTrack.Addr(2))
   
   m_iCDDBID = RShift((n Mod &HFF), 24) Or RShift(t, 8) Or numTracks
   Debug.Print Hex(m_iCDDBID)

End Sub

Private Function CDDBSum(ByVal n As Long) As Long
Dim retVal As Long
   Do While (n > 0)
      retVal = UnsignedAdd(retVal, (n Mod 10))
      n = n \ 10
   Loop
   CDDBSum = retVal
End Function



Private Function UnsignedAdd(Start As Long, Incr As Long) As Long
' This function is useful when doing pointer arithmetic,
' but note it only works for positive values of Incr

   If Start And &H80000000 Then 'Start < 0
      UnsignedAdd = Start + Incr
   ElseIf (Start Or &H80000000) < -Incr Then
      UnsignedAdd = Start + Incr
   Else
      UnsignedAdd = (Start + &H80000000) + (Incr + &H80000000)
   End If
   
End Function



Private Function RShift(ByVal lThis As Long, ByVal lBits As Long) As Long
   If (lBits <= 0) Then
      RShift = lThis
   ElseIf (lBits > 63) Then
      ' .. error ...
   ElseIf (lBits > 31) Then
      RShift = 0
   Else
      If (lThis And m_lPower2(31 - lBits)) = m_lPower2(31 - lBits) Then
         RShift = (lThis And (m_lPower2(31 - lBits) - 1)) * m_lPower2(lBits) Or m_lPower2(31)
      Else
         RShift = (lThis And (m_lPower2(31 - lBits) - 1)) * m_lPower2(lBits)
      End If
   End If
End Function

Private Function LShift(ByVal lThis As Long, ByVal lBits As Long) As Long
   If (lBits <= 0) Then
      LShift = lThis
   ElseIf (lBits > 63) Then
      ' ... error ...
   ElseIf (lBits > 31) Then
      LShift = 0
   Else
      If (lThis And m_lPower2(31)) = m_lPower2(31) Then
         LShift = (lThis And &H7FFFFFFF) \ m_lPower2(lBits) Or m_lPower2(31 - lBits)
      Else
         LShift = lThis \ m_lPower2(lBits)
      End If
   End If
End Function

Private Sub Init()
   m_lPower2(0) = &H1&
   m_lPower2(1) = &H2&
   m_lPower2(2) = &H4&
   m_lPower2(3) = &H8&
   m_lPower2(4) = &H10&
   m_lPower2(5) = &H20&
   m_lPower2(6) = &H40&
   m_lPower2(7) = &H80&
   m_lPower2(8) = &H100&
   m_lPower2(9) = &H200&
   m_lPower2(10) = &H400&
   m_lPower2(11) = &H800&
   m_lPower2(12) = &H1000&
   m_lPower2(13) = &H2000&
   m_lPower2(14) = &H4000&
   m_lPower2(15) = &H8000&
   m_lPower2(16) = &H10000
   m_lPower2(17) = &H20000
   m_lPower2(18) = &H40000
   m_lPower2(19) = &H80000
   m_lPower2(20) = &H100000
   m_lPower2(21) = &H200000
   m_lPower2(22) = &H400000
   m_lPower2(23) = &H800000
   m_lPower2(24) = &H1000000
   m_lPower2(25) = &H2000000
   m_lPower2(26) = &H4000000
   m_lPower2(27) = &H8000000
   m_lPower2(28) = &H10000000
   m_lPower2(29) = &H20000000
   m_lPower2(30) = &H40000000
   m_lPower2(31) = &H80000000
End Sub

Private Sub Class_Initialize()
   Init
   m_bNoCD = True
End Sub
