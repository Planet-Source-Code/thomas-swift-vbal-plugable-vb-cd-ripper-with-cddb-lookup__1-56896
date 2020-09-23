VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFreeDbReadResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum FreeDbReadReturnCode
   RRReadOk = 210  'OK, CDDB database entry follows (until terminating marker)
   RRNotFound = 401   'Specified CDDB entry not found.
   RRServerError = 402   'Server error.
   RRDatabaseEntryCorrupt = 403   'Database entry is corrupt.
   RRNoHandshake = 409  'No handshake.
   RRUnspecifiedError = 500
End Enum

Private m_iCode As Long
Private m_sComments As String
Private m_sPlayOrder As String
Private m_iTrackCount As Long
Private m_sTitles() As String
Private m_sExtended() As String
Private m_sExtendedDisk As String

Public Property Get TrackCount() As Long
   TrackCount = m_iTrackCount
End Property
Public Property Get Title(ByVal trackIndex As Long) As String
   Title = m_sTitles(trackIndex)
End Property
Public Property Get ExtendedInfo(ByVal trackIndex As Long) As String
   ExtendedInfo = m_sExtended(trackIndex)
End Property
Public Property Get Comments() As String
   Comments = m_sComments
End Property
Public Property Get PlayOrder() As String
   PlayOrder = m_sPlayOrder
End Property
Public Property Get ExtendedDiskInfo() As String
   ExtendedDiskInfo = m_sExtendedDisk
End Property

Public Property Get ReturnCode() As FreeDbReadReturnCode
   ReturnCode = m_iCode
End Property
Public Property Get ReturnCodeDescription() As String
   Select Case m_iCode
   Case 210
      ReturnCodeDescription = "CDDB Database Entry Found"
   Case 401
      ReturnCodeDescription = "No CDDB Entry"
   Case 402
      ReturnCodeDescription = "Server Error"
   Case 403
      ReturnCodeDescription = "Database entry is corrupt"
   Case 409
      ReturnCodeDescription = "CDDB Handshaking Error; Check user name and host name"
   Case Else
      ReturnCodeDescription = "Unspecified error"
   End Select
End Property

Friend Sub fInit(ByVal sResponse As String)
   
'<- code categ discid
'<- # xmcd 2.0 CD database file
'<- # ...
'<- (CDDB data...)
'<- .
'   or
'<- code categ discid No such CD entry in database
   
   m_sComments = ""
   m_sPlayOrder = ""
   m_iTrackCount = 0
   Erase m_sTitles()
   Erase m_sExtended()
   m_sExtendedDisk = ""
   m_iCode = 500

   Dim iPos As Long
   iPos = InStr(sResponse, " ")
   
   If (iPos > 0) Then
      ' check the first value:
      If IsNumeric(Left(sResponse, iPos - 1)) Then
         m_iCode = CLng(Left(sResponse, iPos - 1))
         Select Case m_iCode
         Case 210
            ' ok process tracks
            Dim i As Long
            Dim iLineCount As Long
            Dim sLine() As String
            Dim bComplete As Boolean
            SplitDelimitedString sResponse, vbCr, sLine(), iLineCount
            For i = 1 To iLineCount
               ProcessLine sLine(i), bComplete
               If (bComplete) Then
                  Exit For
               End If
            Next i
         Case Else
            ' nothing to do
         End Select
      End If
   End If
   
End Sub

Private Sub ProcessLine( _
      ByVal sResponse As String, _
      ByRef bComplete As Boolean _
   )
   
   Dim iPos As Long
   
   ' Get rid of any line feeds
   iPos = 1
   Do
      iPos = InStr(iPos, sResponse, vbLf)
      If (iPos > 0) Then
         If (iPos = 1) Then
            sResponse = Mid(sResponse, 2)
         ElseIf (iPos < Len(sResponse)) Then
            sResponse = Left(sResponse, iPos - 1) & Mid(sResponse, iPos + 1)
         Else
            sResponse = Left(sResponse, iPos - 1)
         End If
      End If
   Loop While iPos > 0
   
   If (Left(sResponse, 1) = "#") Then
      ' line is a comment
      If Len(m_sComments) > 0 Then
         m_sComments = m_sComments & vbCrLf
      End If
      m_sComments = m_sComments & Mid(sResponse, 2)
   ElseIf (Left(sResponse, 1) = ".") Then
      ' complete
      bComplete = True
   Else
      iPos = InStr(sResponse, "=")
      If (iPos > 1) Then
         Dim sToken As String
         sToken = Left(sResponse, iPos - 1)
         Select Case True
         Case InStr(sToken, "TTITLE") > 0
            AddTitle sResponse
         Case InStr(sToken, "EXTT") > 0
            AddExtendedDetails sResponse
         Case InStr(sToken, "EXTD") > 0
            m_sExtendedDisk = Trim(Mid(sResponse, iPos + 1))
         Case InStr(sToken, "PLAYORDER") > 0
            m_sPlayOrder = Trim(Mid(sResponse, iPos + 1))
         End Select
      End If
   End If
End Sub
Private Sub AddTitle( _
      ByVal sLine As String _
   )
Dim iPos As Long
Dim iNextPos As Long
Dim sTrack As String
Dim lTrack As Long
   iPos = InStr(sLine, "TTITLE")
   iNextPos = InStr(iPos + 6, sLine, "=")
   If (iNextPos > iPos) Then
      sTrack = Mid(sLine, iPos + 6, iNextPos - iPos - 6)
      If IsNumeric(sTrack) Then
         lTrack = CLng(sTrack) + 1
         If (m_iTrackCount < lTrack) Then
            m_iTrackCount = lTrack
            ReDim Preserve m_sTitles(1 To m_iTrackCount) As String
            ReDim Preserve m_sExtended(1 To m_iTrackCount) As String
         End If
         m_sTitles(lTrack) = Trim(Mid(sLine, iNextPos + 1))
      End If
   End If
End Sub
Private Sub AddExtendedDetails( _
      ByVal sLine As String _
   )
Dim iPos As Long
Dim iNextPos As Long
Dim sTrack As String
Dim lTrack As Long
   iPos = InStr(sLine, "EXTT")
   iNextPos = InStr(iPos + 4, sLine, "=")
   If (iNextPos > iPos) Then
      sTrack = Mid(sLine, iPos + 4, iNextPos - iPos - 4)
      If IsNumeric(sTrack) Then
         lTrack = CLng(sTrack) + 1
         If (m_iTrackCount < lTrack) Then
            m_iTrackCount = lTrack
            ReDim Preserve m_sTitles(1 To m_iTrackCount) As String
            ReDim Preserve m_sExtended(1 To m_iTrackCount) As String
         End If
         m_sExtended(lTrack) = Trim(Mid(sLine, iNextPos + 1))
      End If
   End If

End Sub

Private Sub SplitDelimitedString( _
        ByVal sString As String, _
        ByVal sDelim As String, _
        ByRef sValues() As String, _
        ByRef iCount As Long _
    )
' ==================================================================
' Splits sString into an array of parts which are
' delimited in the string by sDelim.  The array is
' indexed 1-iCount where iCount is the number of
' items.  If no items found iCount=1 and the array has
' one element, the original string.
'   sString : String to split
'   sDelim  : Delimiter
'   sValues : Return array of values
'   iCount  : Number of items returned in sValues()
' ==================================================================
Dim iPos As Integer
Dim iNextPos As Integer
Dim iDelimLen As Integer
    iCount = 0
    Erase sValues
    iDelimLen = Len(sDelim)
    iPos = 1
    iNextPos = InStr(sString, sDelim)
    Do While iNextPos > 0
        iCount = iCount + 1
        ReDim Preserve sValues(1 To iCount) As String
        sValues(iCount) = Mid$(sString, iPos, (iNextPos - iPos))
        iPos = iNextPos + iDelimLen
        iNextPos = InStr(iPos, sString, sDelim)
    Loop
    iCount = iCount + 1
    ReDim Preserve sValues(1 To iCount) As String
    sValues(iCount) = Mid$(sString, iPos)
End Sub



