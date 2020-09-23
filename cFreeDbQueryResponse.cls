VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFreeDbQueryResponse"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum FreeDbQueryReturnCode
   QRExactMatch = 200
   QRMoreThanOneMatch = 211
   QRNoMatch = 202
   QRDatabaseEntryCorrupt = 403
   QRCDDBHandshakeError = 409
   QRUnspecifiedError = 500
End Enum

Private Type tFreeDBMatch
   sCategory As String
   sDiscId As String
   sTitle As String
End Type

Private m_iCode As Long

Private m_iMatchCount As Long
Private m_tMatch() As tFreeDBMatch

Public Property Get MatchCount() As Long
   MatchCount = m_iMatchCount
End Property
Public Property Get Category(ByVal matchIndex As Long) As String
   Category = m_tMatch(matchIndex).sCategory
End Property
Public Property Get DiscID(ByVal matchIndex As Long) As String
   DiscID = m_tMatch(matchIndex).sDiscId
End Property
Public Property Get Title(ByVal matchIndex As Long) As String
Dim iPos As Long
   iPos = InStr(m_tMatch(matchIndex).sTitle, " / ")
   If (iPos > 0) Then
      Title = Mid(m_tMatch(matchIndex).sTitle, iPos + 3)
   Else
      Title = m_tMatch(matchIndex).sTitle
   End If
End Property
Public Property Get Artist(ByVal matchIndex As Long) As String
Dim iPos As Long
   iPos = InStr(m_tMatch(matchIndex).sTitle, " / ")
   If (iPos > 0) Then
      Artist = Left(m_tMatch(matchIndex).sTitle, iPos - 1)
   End If
End Property
Public Property Get ReturnCode() As FreeDbQueryReturnCode
   ReturnCode = m_iCode
End Property
Public Property Get ReturnCodeDescription() As String
   Select Case m_iCode
   Case 200
      ReturnCodeDescription = "Exact Match found"
   Case 211
      ReturnCodeDescription = "More than one match found"
   Case 202
      ReturnCodeDescription = "No Matches found"
   Case 403
      ReturnCodeDescription = "Database entry is corrupt"
   Case 409
      ReturnCodeDescription = "CDDB Handshaking Error; Check user name and host name"
   Case Else
      ReturnCodeDescription = "Unspecified error"
   End Select
End Property


Friend Sub fInit(ByVal sResponse As String)
   ' response:
   ' code categ discid dtitle
   ' or
   ' code close matches found
   ' categ discid dtitle
   ' ...
   
   Dim sValues() As String
   Dim iCount As Long
   Dim lCode As String
   Dim iPos As Long
   Dim iNextPos As Long
   Dim i As Long
   
   m_iMatchCount = 0
   Erase m_tMatch
   m_iCode = 500
      
   iPos = InStr(sResponse, " ")
   
   If (iPos > 0) Then
      ' check the first value:
      If IsNumeric(Left(sResponse, iPos - 1)) Then
         m_iCode = CLng(Left(sResponse, iPos - 1))
         Select Case m_iCode
         Case 200
            AddMatch Mid(sResponse, iPos + 1)
            
         Case 211
            ' inexact matches
            iCount = 0
            Erase sValues
            SplitDelimitedString sResponse, vbCr, sValues(), iCount
            For i = 2 To iCount
               AddMatch sValues(i)
            Next i
            
         Case 202
            ' no match found
         Case 403
            ' database is corrupt
         Case 409
            ' no handshake
         Case Else
            ' unknown error
         End Select
      End If
   Else
      ' incorrect response
   End If

End Sub

Private Sub AddMatch(ByVal sResponse As String)
Dim iPos As Long
Dim iNextPos As Long
Dim iState As Long

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

   iPos = 1
   iState = 0
   Do
      iNextPos = InStr(iPos, sResponse, " ")
      If (iNextPos > 0) Then
         Select Case iState
         Case 0
            m_iMatchCount = m_iMatchCount + 1
            ReDim Preserve m_tMatch(1 To m_iMatchCount) As tFreeDBMatch
            m_tMatch(m_iMatchCount).sCategory = Mid(sResponse, iPos, iNextPos - iPos)
         Case 1
            m_tMatch(m_iMatchCount).sDiscId = Mid(sResponse, iPos, iNextPos - iPos)
         Case 2
            m_tMatch(m_iMatchCount).sTitle = Mid(sResponse, iPos)
         End Select
         iState = iState + 1
         iPos = iNextPos + 1
      End If
   Loop While (iNextPos > 0) And (iState < 3)
   
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


