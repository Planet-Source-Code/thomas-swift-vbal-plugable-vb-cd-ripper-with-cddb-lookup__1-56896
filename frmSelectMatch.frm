VERSION 5.00
Begin VB.Form frmSelectMatch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "More Than One Matching CD Found:"
   ClientHeight    =   3030
   ClientLeft      =   5250
   ClientTop       =   5145
   ClientWidth     =   5640
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
   ScaleHeight     =   3030
   ScaleWidth      =   5640
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   435
      Left            =   2640
      TabIndex        =   3
      Top             =   2460
      Width           =   1395
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   4140
      TabIndex        =   2
      Top             =   2460
      Width           =   1395
   End
   Begin VB.ListBox lstMatches 
      Height          =   1620
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   5415
   End
   Begin VB.Label lblInfo 
      Caption         =   "More than one entry has been found for your CD.  Select the one which most closely matches:"
      Height          =   555
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5415
   End
End
Attribute VB_Name = "frmSelectMatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_iSelectedMatch As Long
Private m_cQR As cFreeDbQueryResponse

Public Property Let QueryResponse(cQ As cFreeDbQueryResponse)
   Set m_cQR = cQ
End Property
Public Property Get SelectedMatch() As Long
   SelectedMatch = m_iSelectedMatch
End Property

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   m_iSelectedMatch = lstMatches.ListIndex
   Unload Me
End Sub

Private Sub Form_Load()
   Dim i As Long
   For i = 1 To m_cQR.MatchCount
      lstMatches.AddItem m_cQR.Category(i) & ": " & m_cQR.Artist(i) & " " & m_cQR.Title(i)
   Next i
   lstMatches.ListIndex = 0
End Sub
