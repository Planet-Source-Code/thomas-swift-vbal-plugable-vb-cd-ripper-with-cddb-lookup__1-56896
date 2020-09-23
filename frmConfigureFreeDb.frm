VERSION 5.00
Begin VB.Form frmConfigureFreeDb 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configure FreeDB"
   ClientHeight    =   7350
   ClientLeft      =   3600
   ClientTop       =   3135
   ClientWidth     =   6495
   Icon            =   "frmConfigureFreeDb.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   6495
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3540
      TabIndex        =   21
      Top             =   6840
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4980
      TabIndex        =   20
      Top             =   6840
      Width           =   1335
   End
   Begin VB.PictureBox picTab 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6675
      Left            =   60
      ScaleHeight     =   6675
      ScaleWidth      =   6375
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   6375
      Begin VB.TextBox txtServer 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Text            =   "http://freedb.freedb.org/"
         Top             =   420
         Width           =   4935
      End
      Begin VB.TextBox txtQueryURL 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Text            =   "~cddb/cddb.cgi"
         Top             =   780
         Width           =   4935
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   7
         Text            =   "steve"
         Top             =   1440
         Width           =   4935
      End
      Begin VB.TextBox txtUserHost 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   9
         Text            =   "vbaccelerator.com"
         Top             =   1800
         Width           =   4935
      End
      Begin VB.TextBox txtAgentName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   11
         Text            =   "VBALTrackList"
         Top             =   2340
         Width           =   4935
      End
      Begin VB.TextBox txtAgentVersion 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         TabIndex        =   13
         Text            =   "1.0"
         Top             =   2700
         Width           =   4935
      End
      Begin VB.ComboBox cboTestType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3960
         Width           =   4935
      End
      Begin VB.CommandButton cmdTest 
         Caption         =   "&Test"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1320
         TabIndex        =   17
         Top             =   4380
         Width           =   1275
      End
      Begin VB.TextBox txtResponse 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Left            =   1320
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   19
         Top             =   4920
         Width           =   4935
      End
      Begin VB.Label lblFreeDbServer 
         Caption         =   "&Server:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   420
         Width           =   1155
      End
      Begin VB.Label lblQueryCGI 
         Caption         =   "&Query URL:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   840
         Width           =   1155
      End
      Begin VB.Label lblUserName 
         Caption         =   "User &Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   6
         Top             =   1500
         Width           =   1155
      End
      Begin VB.Label Label1 
         Caption         =   "User &Host:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   8
         Top             =   1860
         Width           =   1155
      End
      Begin VB.Label lblAgentName 
         Caption         =   "&Agent Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   10
         Top             =   2400
         Width           =   1155
      End
      Begin VB.Label lblAgentVersion 
         Caption         =   "Agent &Version"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   60
         TabIndex        =   12
         Top             =   2760
         Width           =   1155
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H80000010&
         Caption         =   " FreeDB Configuration:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   6195
      End
      Begin VB.Label lblTest 
         BackColor       =   &H80000010&
         Caption         =   " Test FreeDB Configuration:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   255
         Left            =   60
         TabIndex        =   14
         Top             =   3660
         Width           =   6195
      End
      Begin VB.Label lblTestType 
         Caption         =   "&Test:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   4020
         Width           =   1155
      End
      Begin VB.Label lblResponse 
         Caption         =   "&Response:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4980
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmConfigureFreeDb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cFreeDb As cFreeDB
Private WithEvents m_cTestFreeDb As cFreeDB
Attribute m_cTestFreeDb.VB_VarHelpID = -1
Private m_bCancelled As Boolean

Private Function SaveConfiguration(cFDB As cFreeDB) As Boolean
   With cFDB
      .Server = txtServer.Text
      .Url = txtQueryURL.Text
      .UserName = txtUserName.Text
      .UserHost = txtUserHost.Text
      .AgentName = txtAgentName.Text
      .AgentVersion = txtAgentVersion.Text
   End With
   SaveConfiguration = True
End Function

Public Property Let FreeDB(cFDB As cFreeDB)
   Set m_cFreeDb = cFDB
End Property

Public Property Get Cancelled() As Boolean
   m_bCancelled = True
End Property

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   If (SaveConfiguration(m_cFreeDb)) Then
      m_bCancelled = False
      Unload Me
   End If
End Sub

Private Sub cmdTest_Click()
   If SaveConfiguration(m_cTestFreeDb) Then
      cmdTest.Enabled = False
      Select Case cboTestType.ListIndex
      Case 0
         m_cTestFreeDb.Command = "stat"
      Case 1
         m_cTestFreeDb.Command = "ver"
      Case 2
         m_cTestFreeDb.Command = "motd"
      Case 3
         m_cTestFreeDb.Command = "cddb lscat"
      End Select
      m_cTestFreeDb.Start
   End If
End Sub

Private Sub Form_Load()
   m_bCancelled = True
   cboTestType.AddItem "Status"
   cboTestType.AddItem "Version"
   cboTestType.AddItem "Message of the Day"
   cboTestType.AddItem "Category List"
   cboTestType.ListIndex = 0
   Set m_cTestFreeDb = New cFreeDB
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ' in case the test is still running:
   m_cTestFreeDb.Abort
   ' clear up:
   Set m_cTestFreeDb = Nothing
End Sub

Private Sub m_cTestFreeDb_CommandReady()
   cmdTest.Enabled = True
   txtResponse.Text = m_cTestFreeDb.Response
End Sub
