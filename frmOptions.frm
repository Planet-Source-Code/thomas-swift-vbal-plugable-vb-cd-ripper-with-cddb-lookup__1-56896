VERSION 5.00
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "VB CD Ripper Options"
   ClientHeight    =   5730
   ClientLeft      =   5070
   ClientTop       =   2445
   ClientWidth     =   6090
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
   ScaleHeight     =   5730
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picRipper 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   180
      ScaleHeight     =   4575
      ScaleWidth      =   5655
      TabIndex        =   2
      Top             =   240
      Width           =   5655
      Begin VB.ComboBox cboDrives 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   0
         Width           =   4215
      End
      Begin VB.TextBox txtReadSectors 
         Height          =   315
         Left            =   1440
         TabIndex        =   11
         Top             =   480
         Width           =   1395
      End
      Begin VB.TextBox txtReadOverlap 
         Height          =   315
         Left            =   1440
         TabIndex        =   10
         Top             =   840
         Width           =   1395
      End
      Begin VB.TextBox txtStartOffset 
         Height          =   315
         Left            =   1440
         TabIndex        =   9
         Top             =   1200
         Width           =   1395
      End
      Begin VB.TextBox txtEndOffset 
         Height          =   315
         Left            =   1440
         TabIndex        =   8
         Top             =   1560
         Width           =   1395
      End
      Begin VB.TextBox txtRetries 
         Height          =   315
         Left            =   4260
         TabIndex        =   7
         Top             =   480
         Width           =   1395
      End
      Begin VB.TextBox txtBlockCompare 
         Height          =   315
         Left            =   4260
         TabIndex        =   6
         Top             =   840
         Width           =   1395
      End
      Begin VB.TextBox txtCDSpeed 
         Height          =   315
         Left            =   4260
         TabIndex        =   5
         Top             =   1200
         Width           =   1395
      End
      Begin VB.TextBox txtSpinUpTime 
         Height          =   315
         Left            =   4260
         TabIndex        =   4
         Top             =   1560
         Width           =   1395
      End
      Begin VB.ComboBox cboDriveType 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2160
         Width           =   4215
      End
      Begin VB.Label lblDrives 
         Caption         =   "&Drive:"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   60
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Read &Sectors:"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   540
         Width           =   1305
      End
      Begin VB.Label Label2 
         Caption         =   "Read O&verlap:"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   900
         Width           =   1305
      End
      Begin VB.Label Label3 
         Caption         =   "S&tart Offset:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1260
         Width           =   1305
      End
      Begin VB.Label Label4 
         Caption         =   "&End Offset:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   1620
         Width           =   1305
      End
      Begin VB.Label Label5 
         Caption         =   "Retr&ies:"
         Height          =   255
         Left            =   2940
         TabIndex        =   17
         Top             =   540
         Width           =   1305
      End
      Begin VB.Label Label6 
         Caption         =   "&Block Compare:"
         Height          =   255
         Left            =   2940
         TabIndex        =   16
         Top             =   900
         Width           =   1305
      End
      Begin VB.Label Label7 
         Caption         =   "&CD Speed:"
         Height          =   255
         Left            =   2940
         TabIndex        =   15
         Top             =   1260
         Width           =   1305
      End
      Begin VB.Label Label8 
         Caption         =   "Spin &Up Time (s):"
         Height          =   255
         Left            =   2940
         TabIndex        =   14
         Top             =   1620
         Width           =   1305
      End
      Begin VB.Label Label9 
         Caption         =   "Drive T&ype:"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   1335
      End
   End
   Begin VB.PictureBox picPlugin 
      BorderStyle     =   0  'None
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4575
      ScaleWidth      =   5655
      TabIndex        =   23
      Top             =   780
      Width           =   5655
      Begin VB.CommandButton cmdAddNew 
         Caption         =   "&Add..."
         Height          =   315
         Left            =   4800
         TabIndex        =   32
         Top             =   0
         Width           =   795
      End
      Begin VB.ComboBox cboParam 
         Height          =   315
         Index           =   0
         Left            =   1500
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   2100
         Visible         =   0   'False
         Width           =   4155
      End
      Begin VB.TextBox txtPluginInfo 
         Height          =   1275
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   26
         Top             =   420
         Width           =   4215
      End
      Begin VB.ComboBox cboPlugin 
         Height          =   315
         Left            =   1440
         Style           =   2  'Dropdown List
         TabIndex        =   24
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label lblConfig 
         Caption         =   "*param*"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   2160
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblConfiguration 
         BackColor       =   &H80000010&
         Caption         =   " Configuration"
         ForeColor       =   &H80000014&
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   1800
         Width           =   5535
      End
      Begin VB.Label lblInfo 
         Caption         =   "&Info:"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblPlugin 
         Caption         =   "Output &Plugin:"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   60
         Width           =   1335
      End
   End
   Begin vbalDTab6.vbalDTabControl tabOptions 
      Height          =   5055
      Left            =   60
      TabIndex        =   31
      Top             =   60
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   8916
      TabAlign        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   3180
      TabIndex        =   0
      Top             =   5220
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   435
      Left            =   4560
      TabIndex        =   1
      Top             =   5220
      Width           =   1335
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cRip As cCDRip
Private m_iLastIndex As Long

Private Sub ShowPluginDetails(ByVal nIndex As Long)
Dim info As IPluginInformation
   Set info = PluginManagerInstance.PluginInformationByIndex(nIndex)
   If Not (info Is Nothing) Then
      Dim sText As String
      sText = "Description:" & vbCrLf & info.PluginDescription("en-us")
      sText = sText & vbCrLf & vbCrLf
      sText = sText & "Author:" & info.PluginAuthor("en-us") & vbCrLf
      sText = sText & "Website:" & info.PluginWebsite("en-us") & vbCrLf
      sText = sText & vbCrLf & vbCrLf
      Dim sAck As String
      sAck = info.PluginAcknowledgements("en-us")
      If Len(sAck) = 0 Then sAck = "(none)"
      sText = sText & "Acknowledgments:" & vbCrLf & sAck
      txtPluginInfo.Text = sText
      
      Dim config As IPluginConfig
      Set config = info.Configuration
      If Not (config Is Nothing) Then
         '
      Else
         '
      End If
   Else
      txtPluginInfo.Text = "No information available for this plugin."
   End If
   PluginManagerInstance.SelectedPluginIndex = nIndex
End Sub

Private Sub ShowPlugins()
Dim i As Long
Dim info As IPluginInformation
   
   cboPlugin.Clear

   With PluginManagerInstance
      For i = 1 To .PluginCount
         Set info = .PluginInformationByIndex(i)
         If Not (info Is Nothing) Then
            cboPlugin.AddItem info.PluginName("en-us")
         Else
            cboPlugin.AddItem .PluginProgId(i)
         End If
      Next i
   
      If (cboPlugin.ListCount > 0) Then
         cboPlugin.ListIndex = .SelectedPluginIndex - 1
      End If
      
   End With

End Sub

Public Property Let CRRip(cRip As cCDRip)
   Set m_cRip = cRip
End Property

Private Sub ShowDrives()
Dim i As Long
   
   For i = 1 To m_cRip.CDDriveCount
      cboDrives.AddItem m_cRip.CDDrive(i).Name
   Next i
   
   If (cboDrives.ListCount > 0) Then
      cboDrives.ListIndex = 0
   Else
      Dim ctl As Control
      Dim txt As TextBox
      Dim cbo As ComboBox
      For Each ctl In Me.Controls
         If (TypeOf ctl Is TextBox) Then
            Set txt = ctl
            txt.Enabled = False
            txt.Text = ""
            txt.BackColor = vbButtonFace
         ElseIf (TypeOf ctl Is ComboBox) Then
            Set cbo = ctl
            cbo.Enabled = False
            cbo.ListIndex = -1
            cbo.BackColor = vbButtonFace
         End If
      Next
      cmdOK.Enabled = False
   End If
   
End Sub

Private Sub cboDrives_Click()
   '
   If (m_iLastIndex > 0) And (m_iLastIndex <> cboDrives.ListIndex + 1) Then
      m_cRip.CDDrive(m_iLastIndex).Apply
   End If
   
   Dim cD As cDrive
   Set cD = m_cRip.CDDrive(cboDrives.ListIndex + 1)
   
   txtReadSectors.Text = cD.ReadSectors
   txtReadOverlap.Text = cD.ReadOverlap
   txtStartOffset.Text = cD.StartOffset
   txtEndOffset.Text = cD.EndOffset
   txtRetries.Text = cD.Retries
   txtBlockCompare.Text = cD.BlockCompare
   txtCDSpeed.Text = cD.CDSpeed
   txtSpinUpTime.Text = cD.SpinUpTime
   cboDriveType.ListIndex = cD.DriveType
   
   m_iLastIndex = cboDrives.ListIndex + 1
   
   '
End Sub

Private Sub cboPlugin_Click()
   ShowPluginDetails cboPlugin.ListIndex + 1
End Sub

Private Sub cmdAddNew_Click()
Dim sProgId As String
   sProgId = InputBox("Enter the ProgID of the Plugin:", "Add New Plugin", "")
   sProgId = Trim(sProgId)
   If Len(sProgId) > 0 Then
      If (PluginManagerInstance.AddPlugin(sProgId)) Then
         ShowPlugins
      End If
   End If
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdOK_Click()
   m_cRip.CDDrive(cboDrives.ListIndex + 1).Apply
   Unload Me
End Sub

Private Sub Form_Load()
Dim cTab As cTab

   tabOptions.ShowCloseButton = False
   tabOptions.AllowScroll = False
   Set cTab = tabOptions.Tabs.Add("PLUGINS", , "Plugins")
   cTab.Panel = picPlugin
   Set cTab = tabOptions.Tabs.Add("RIPPER", , "Ripper")
   cTab.Panel = picRipper
   
   cboDriveType.AddItem "GENERIC"
   cboDriveType.AddItem "TOSHIBA"
   cboDriveType.AddItem "TOSHIBANEW"
   cboDriveType.AddItem "IBM"
   cboDriveType.AddItem "NEC"
   cboDriveType.AddItem "DEC"
   cboDriveType.AddItem "IMS"
   cboDriveType.AddItem "KODAK"
   cboDriveType.AddItem "RICOH"
   cboDriveType.AddItem "HP"
   cboDriveType.AddItem "PHILIPS"
   cboDriveType.AddItem "PLASMON"
   cboDriveType.AddItem "GRUNDIGCDR100IPW"
   cboDriveType.AddItem "MITSUMICDR"
   cboDriveType.AddItem "PLEXTOR"
   cboDriveType.AddItem "SONY"
   cboDriveType.AddItem "YAMAHA"
   cboDriveType.AddItem "NRC"
   cboDriveType.AddItem "IMSCDD5"
   cboDriveType.AddItem "CUSTOMDRIVE"
   
   ShowDrives
   
   ShowPlugins
   
End Sub

