VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTrack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_iTrack As Long

Private m_sTitle As String
Private m_sAlbum As String
Private m_sArtist As String
Private m_sAlbumArtist As String
Private m_sTrack As String
Private m_sYear As String
Private m_sComment As String
Private m_sFileName As String

Public Sub Init(ByVal iTrack As Long, _
      ByVal sTitle As String, _
      ByVal sAlbum As String, _
      ByVal sAlbumArtist As String, _
      ByVal sArtist As String, _
      ByVal sTrack As String, _
      ByVal sYear As String, _
      ByVal sComment As String _
   )
   m_iTrack = iTrack
   m_sTitle = sTitle
   m_sAlbum = sAlbum
   m_sAlbumArtist = sAlbumArtist
   m_sArtist = sArtist
   m_sTrack = sTrack
   m_sYear = sYear
   m_sComment = sComment
   
   DeriveFileName
   
End Sub

Public Property Get FileName() As String
   FileName = m_sFileName
End Property

Private Sub DeriveFileName()
Dim sArtist As String
Dim sAlbumArtist As String
Dim sAlbum As String
Dim sTitle As String

   m_sFileName = ""

   sAlbumArtist = Trim(StripBadChars(m_sAlbumArtist))
   If Len(sAlbumArtist) > 0 Then
      m_sFileName = sAlbumArtist
   End If
   
   sAlbum = Trim(StripBadChars(m_sAlbum))
   If (Len(sAlbum) > 0) Then
      If Len(m_sFileName) > 0 Then
         m_sFileName = m_sFileName & "\"
      End If
      m_sFileName = m_sFileName & sAlbum
   End If
   
   sTitle = Trim(StripBadChars(m_sTitle))
   If Len(sTitle) = 0 Then
      sTitle = "Track " & Format(m_iTrack, "##0")
   End If
   
   sArtist = Trim(StripBadChars(m_sArtist))
   If Not (StrComp(sArtist, sAlbumArtist, vbTextCompare) = 0) Then
      sTitle = sArtist & " - " & sTitle
   End If
   
   
   If Len(m_sFileName) > 0 Then
      m_sFileName = m_sFileName & "\"
   End If
   m_sFileName = m_sFileName & sTitle
   
End Sub

Private Function StripBadChars(ByVal sString As String) As String
Dim i As Long
Dim sC As String
Dim sRet As String
Dim bLastCharDot As Boolean
   For i = 1 To Len(sString)
      sC = Mid(sString, i, 1)
      Select Case sC
      Case "\", "/", ":", "?", "*", "<", ">", "|"
         sRet = sRet & "_"
         bLastCharDot = False
      Case "."
         If (bLastCharDot) Then
            ' skip it
         Else
            bLastCharDot = True
            sRet = sRet & "."
         End If
      Case Else
         If (Asc(sC) < 32) Then
            sRet = sRet & "_"
         Else
            sRet = sRet & sC
         End If
         bLastCharDot = False
      End Select
   Next i
   If Len(sRet) > 1 Then
      If Right(sRet, 1) = "." Then
         If Len(sRet) = 1 Then
            sRet = ""
         Else
            sRet = Left(sRet, Len(sRet) - 1)
         End If
      End If
   End If
   StripBadChars = sRet
End Function
   
Public Property Get CDTrackNumber() As Long
   CDTrackNumber = m_iTrack
End Property
Public Property Get Title() As String
   Title = m_sTitle
End Property
Public Property Get Album() As String
   Album = m_sAlbum
End Property
Public Property Get Track() As String
   Track = m_sTrack
End Property
Public Property Get Year() As String
   Year = m_sYear
End Property
Public Property Get Artist() As String
   Artist = m_sArtist
End Property
Public Property Get Comment() As String
   Comment = m_sComment
End Property


