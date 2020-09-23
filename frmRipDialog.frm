VERSION 5.00
Begin VB.Form frmRipDialog 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CD Rip in Progress"
   ClientHeight    =   2460
   ClientLeft      =   5820
   ClientTop       =   4410
   ClientWidth     =   5115
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   5115
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picTrack 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   4875
      TabIndex        =   3
      Top             =   1500
      Width           =   4875
   End
   Begin VB.PictureBox picSelected 
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   4875
      TabIndex        =   2
      Top             =   1080
      Width           =   4875
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   1920
      TabIndex        =   0
      Top             =   1920
      Width           =   1275
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      Height          =   855
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "frmRipDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cPrgSelected As cProgressBar
Private m_cPrgTrack As cProgressBar

Private m_cToc As cToc
Private m_selTracks As Collection
Private m_sDir As String
Private m_bRipping As Boolean
Private m_bCancel As Boolean
Private m_writer As IWaveDataWriter

Private Sub EnsureDirectory(ByVal sFile As String)
Dim sParts() As String
Dim iCount As Long
Dim i As Long
Dim sC As String
   
   iCount = iCount + 1
   ReDim Preserve sParts(1 To iCount) As String
   For i = 1 To Len(sFile)
      sC = Mid(sFile, i, 1)
      If (sC = "\") Then
         iCount = iCount + 1
         ReDim Preserve sParts(1 To iCount) As String
      Else
         sParts(iCount) = sParts(iCount) & sC
      End If
   Next i
   
Dim sDir As String
   sDir = sParts(1)
   For i = 2 To iCount - 1
      sDir = sDir & "\" & sParts(i)
      If Not (DirectoryExists(sDir)) Then
         MkDir sDir
      End If
   Next i
      
End Sub

Private Function DirectoryExists(ByVal sDir As String) As Boolean
Dim sChk As String
   On Error Resume Next
   sChk = Dir(sDir, vbDirectory)
   DirectoryExists = ((Err.Number = 0) And Len(sChk) > 0)
End Function

Public Property Let Writer(dataWriter As IWaveDataWriter)
   Set m_writer = dataWriter
End Property

Public Sub RipSelected()
Dim i As Long
   
   m_bRipping = True
   m_cPrgSelected.Max = m_selTracks.Count
   m_cPrgSelected.Value = 0
   For i = 1 To m_selTracks.Count
      If (m_bRipping) Then
         Rip m_selTracks(i)
         m_cPrgSelected.Value = i
      End If
   Next i
   m_bRipping = False
   Unload Me
   
End Sub

Private Sub Rip(ByVal cT As cTrack)
Dim sFile As String
Dim cTrack As cTocEntry
Dim cTrackRip As cCDTrackRipper
Dim lErr As Long
Dim sMsg As String
   
On Error GoTo fileErrorHandler
   m_cPrgTrack.Value = 0
   
   sFile = m_sDir & cT.FileName & "." & m_writer.FileExtension
   EnsureDirectory sFile
   If (m_writer.OpenFile(sFile)) Then
      
On Error GoTo writeErrorHandler
      m_writer.Album = cT.Album
      m_writer.Artist = cT.Artist
      m_writer.Comment = cT.Comment
      m_writer.Title = cT.Title
      m_writer.TrackNumber = cT.Track
      m_writer.Year = cT.Year
      'm_writer.Genre = cT.Genre
      
      Set cTrack = m_cToc.Entry(cT.CDTrackNumber)
      
      lblInfo.Caption = "Ripping track " & cTrack.TrackNumber & _
         " (" & m_cPrgSelected.Value + 1 & " of " & m_cPrgSelected.Max & ") to " & _
         sFile & "..."
      
      Set cTrackRip = New cCDTrackRipper
      cTrackRip.CreateForTrack cTrack
      
      If (cTrackRip.OpenRipper()) Then
         m_cPrgTrack.Max = 100
         m_cPrgTrack.Value = 0
         Do While cTrackRip.Read
            m_writer.WriteWavData cTrackRip.ReadBufferPtr, cTrackRip.ReadBufferSize
            m_cPrgTrack.Value = cTrackRip.PercentComplete
            DoEvents
            If (m_bCancel) Then
               Exit Do
            End If
         Loop
         
writeErrorHandler:
         lErr = Err.Number: sMsg = Err.Description
         m_bCancel = m_bCancel Or (Not (Err.Number = 0))
         
         cTrackRip.CloseRipper
         m_writer.CloseFile
         
         If (m_bCancel) Then
            lblInfo.Caption = "Cancelled"
            On Error Resume Next
            Kill sFile
            If Not (lErr = 0) Then
               MsgBox "An error occurred during ripping; " & sMsg, vbExclamation
            End If
            m_bRipping = False
         Else
            lblInfo.Caption = "Completed track " & cTrack.TrackNumber
         End If
         
      End If
   End If
   Exit Sub
   
fileErrorHandler:
   MsgBox Err.Description, vbInformation
   Exit Sub
End Sub

Public Property Let OutputDir(ByVal sDir As String)
   m_sDir = sDir
   If (Right(sDir, 1) <> "\") Then
      m_sDir = m_sDir & "\"
   End If
End Property

Public Property Let RipTOC(cT As cToc)
   Set m_cToc = cT
End Property

Public Sub AddTrack(cT As cTrack)
   m_selTracks.Add cT, "C" & m_selTracks.Count + 1
End Sub

Private Sub cmdCancel_Click()
   '
   If (m_bRipping) Then
      m_bCancel = True
   Else
      Unload Me
   End If
   '
End Sub

Private Sub Form_Initialize()
   
   Set m_selTracks = New Collection

   Set m_cPrgSelected = New cProgressBar
   m_cPrgSelected.XpStyle = True
   m_cPrgSelected.Value = 0
   m_cPrgSelected.Max = 0
   Set m_cPrgTrack = New cProgressBar
   m_cPrgTrack.XpStyle = True
   m_cPrgTrack.Value = 0
   m_cPrgTrack.Max = 0
   
End Sub

Private Sub Form_Load()
   '
   m_cPrgSelected.DrawObject = picSelected
   m_cPrgSelected.Value = 0
   m_cPrgTrack.DrawObject = picTrack
   m_cPrgTrack.Value = 0
   
   lblInfo.Caption = "Preparing to rip " & m_selTracks.Count & " tracks..."
   
   Me.Show
   Me.Refresh
   '
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   If (m_bRipping) Then
      cmdCancel_Click
      Cancel = True
   End If
End Sub

Private Sub picSelected_Paint()
   m_cPrgSelected.Draw
End Sub

Private Sub picTrack_Paint()
   m_cPrgTrack.Draw
End Sub
