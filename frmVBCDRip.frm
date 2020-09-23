VERSION 5.00
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "VBALSGRID6.OCX"
Begin VB.Form frmVBCDRip 
   Caption         =   "VB CD Ripper"
   ClientHeight    =   5865
   ClientLeft      =   3300
   ClientTop       =   2670
   ClientWidth     =   6705
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmVBCDRip.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5865
   ScaleWidth      =   6705
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   120
      Top             =   5385
   End
   Begin VB.TextBox txtEdit 
      Height          =   315
      Left            =   4785
      TabIndex        =   14
      Top             =   5475
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.PictureBox pnlMain 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   4935
      Left            =   0
      ScaleHeight     =   4935
      ScaleWidth      =   6705
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   6705
      Begin VB.TextBox txtYear 
         Height          =   315
         Left            =   1140
         TabIndex        =   10
         Top             =   1260
         Width           =   2775
      End
      Begin VB.CheckBox chkCompilation 
         Caption         =   "&Compilation"
         Height          =   255
         Left            =   3960
         TabIndex        =   6
         Top             =   540
         Width           =   1815
      End
      Begin VB.TextBox txtArtist 
         Height          =   315
         Left            =   1140
         TabIndex        =   5
         Top             =   540
         Width           =   2775
      End
      Begin VB.TextBox txtAlbum 
         Height          =   315
         Left            =   1140
         TabIndex        =   8
         Top             =   900
         Width           =   2775
      End
      Begin vbAcceleratorSGrid6.vbalGrid grdTracks 
         Height          =   2835
         Left            =   60
         TabIndex        =   11
         Top             =   1620
         Width           =   6555
         _ExtentX        =   11562
         _ExtentY        =   5001
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         DisableIcons    =   -1  'True
      End
      Begin VB.ComboBox cboDrives 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   60
         Width           =   4095
      End
      Begin VB.TextBox txtCDDBQuery 
         Height          =   315
         Left            =   1200
         TabIndex        =   13
         Tag             =   "WAITING"
         Top             =   4560
         Width           =   5295
      End
      Begin VB.CommandButton cmdRefresh 
         Caption         =   "&Refresh"
         Height          =   315
         Left            =   5280
         TabIndex        =   3
         Top             =   60
         Width           =   1155
      End
      Begin VB.Label lblYear 
         Caption         =   "&Year:"
         Height          =   255
         Left            =   60
         TabIndex        =   9
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label lblArtist 
         Caption         =   "Ar&tist:"
         Height          =   255
         Left            =   60
         TabIndex        =   4
         Top             =   600
         Width           =   1035
      End
      Begin VB.Label lblAlbum 
         Caption         =   "&Album"
         Height          =   255
         Left            =   60
         TabIndex        =   7
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label lblDrives 
         Caption         =   "&Drive:"
         Height          =   255
         Left            =   60
         TabIndex        =   1
         Top             =   120
         Width           =   1035
      End
      Begin VB.Label lblCDDBQuery 
         Caption         =   "CDDB &Query:"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   4620
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.Menu mnuFileTOP 
      Caption         =   "&File"
      Begin VB.Menu mnuFile 
         Caption         =   "&Rip..."
         Index           =   0
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   1
      End
      Begin VB.Menu mnuFile 
         Caption         =   "&Configure..."
         Index           =   2
      End
      Begin VB.Menu mnuFile 
         Caption         =   "-"
         Index           =   3
      End
      Begin VB.Menu mnuFile 
         Caption         =   "E&xit"
         Index           =   4
      End
   End
   Begin VB.Menu mnuEditTop 
      Caption         =   "&Edit"
      Begin VB.Menu mnuEdit 
         Caption         =   "&Select All"
         Index           =   0
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Invert Selection"
         Index           =   1
      End
   End
   Begin VB.Menu mnuHelpTOP 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "&About..."
         Index           =   0
      End
   End
End
Attribute VB_Name = "frmVBCDRip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_cRip As cCDRip
Private m_cToc As cToc
Private WithEvents m_cFreeDb As cFreeDB
Attribute m_cFreeDb.VB_VarHelpID = -1

Private Sub ConfigureGrid()
   
   With grdTracks
      .Editable = True
      .SingleClickEdit = True
      .AddColumn "Selected", lColumnWidth:=20
      .AddColumn "Name", "Name"
      .AddColumn "Artist", "Artist"
      .AddColumn "Track", "Track", ecgHdrTextALignCentre, sFmtString:="#0"
      .AddColumn "StartTime", "Start Time", ecgHdrTextALignRight
      .AddColumn "Length", "Length", ecgHdrTextALignRight
      .AddColumn "Size", "Size (Mb)", ecgHdrTextALignRight, sFmtString:="#0.00MB"
      .GridLines = True
      .GridLineMode = ecgGridFillControl
      .GridLineColor = BlendColor(vbHighlight, vbWindowBackground, 96)
      .GridFillLineColor = .GridLineColor
      .NoHorizontalGridLines = True
      .SelectionAlphaBlend = True
      .DrawFocusRectangle = False
      .SelectionOutline = True
      chkCompilation_Click
   End With
   
End Sub

Private Sub EnableControls(ByVal bState As Boolean)
Dim ctl As Control
Dim txt As TextBox
Dim chk As CheckBox
Dim cbo As ComboBox
Dim pic As PictureBox
Dim mnu As Menu
   
   For Each ctl In Controls
      If TypeOf ctl Is TextBox Then
         Set txt = ctl
         txt.Enabled = bState
         txt.BackColor = IIf(bState, vbWindowBackground, vbButtonFace)
      ElseIf TypeOf ctl Is ComboBox Then
         Set cbo = ctl
         cbo.Enabled = bState
         cbo.BackColor = IIf(bState, vbWindowBackground, vbButtonFace)
      ElseIf TypeOf ctl Is Menu Then
         Set mnu = ctl
         If InStr(mnu.Name, "TOP") = 0 Then
            If Trim(mnu.Caption) <> "-" Then
               mnu.Enabled = bState
            End If
         End If
      ElseIf TypeOf ctl Is Label Then
         ' do nothing
      Else
         ctl.Enabled = bState
      End If
   Next
   
   If (bState) Then
      chkCompilation_Click
   End If
   
End Sub

Private Sub ShowDrives()
On Error Resume Next
Dim i As Long
   
   Set m_cRip = New cCDRip
   m_cRip.Create App.Path & "\cdrip.ini" ' this INI file isn't used currently
   
   For i = 1 To m_cRip.CDDriveCount
      cboDrives.AddItem m_cRip.CDDrive(i).Name
   Next i
   
   If (cboDrives.ListCount > 0) Then
      cboDrives.ListIndex = 0
   Else
      ShowTracks
   End If
      
End Sub

Private Sub ShowTracks()
Dim lIndex As Long
Dim sFnt As New StdFont
Dim bEnable As Boolean
Dim sCDDB As String
   
   lIndex = cboDrives.ListIndex + 1
   If (lIndex > 0) Then
      ' Check if anything has changed:
      Dim cD As cDrive
      Set cD = m_cRip.CDDrive(lIndex)
      If (cD.IsUnitReady) Then
         sCDDB = cD.TOC.CDDBQuery
      End If
   End If
   
   If (sCDDB = txtCDDBQuery.Text) Then
      bEnable = (Len(sCDDB) > 0)
   Else
      txtArtist.Tag = ""
      txtArtist.Text = ""
      txtAlbum.Text = ""
      txtYear.Text = ""
   
      sFnt.Name = "Marlett"
      sFnt.Size = 10
      
      grdTracks.Clear
      EnableControls False
         
      txtCDDBQuery.Text = ""
      Set m_cToc = Nothing
   
      If (lIndex > 0) Then
         If (cD.IsUnitReady) Then
            
            Set m_cToc = cD.TOC
            txtCDDBQuery.Text = m_cToc.CDDBQuery
            
            Dim i As Long
            For i = 1 To m_cToc.Count
               grdTracks.AddRow
               With m_cToc.Entry(i)
                  grdTracks.CellDetails i, 1, "b", oFont:=sFnt, lItemData:=True
                  grdTracks.CellText(i, 2) = "Track " & Format(.TrackNumber, "00")
                  grdTracks.CellText(i, 3) = ""
                  grdTracks.CellDetails i, 4, .TrackNumber, DT_SINGLELINE Or DT_VCENTER Or DT_CENTER, oForeColor:=&H808080
                  grdTracks.CellDetails i, 5, .FormattedStartTime, DT_SINGLELINE Or DT_VCENTER Or DT_CENTER, oForeColor:=&H808080
                  grdTracks.CellDetails i, 6, .FormattedLength, DT_SINGLELINE Or DT_VCENTER Or DT_CENTER, oForeColor:=&H808080
                  grdTracks.CellDetails i, 7, .SizeBytes / (1024& * 1024&), DT_SINGLELINE Or DT_VCENTER Or DT_CENTER, oForeColor:=&H808080
               End With
               SelectRow i, True
            Next i
            
            bEnable = (m_cToc.Count > 0)
            
         Else
            grdTracks.AddRow
            grdTracks.CellText(1, 2) = "No CD In Drive"
         End If
      Else
         grdTracks.AddRow
         grdTracks.CellText(1, 2) = "No CD Selected"
      End If
   End If
   
   EnableControls bEnable

   If Not (bEnable) Then
      ' Make sure can still access stuff
      pnlMain.Enabled = True
      cboDrives.Enabled = True
      cboDrives.BackColor = vbWindowBackground
      cmdRefresh.Enabled = True
      mnuFile(2).Enabled = True
      mnuHelp(0).Enabled = True
   Else
      If (grdTracks.Rows > 0) Then
         grdTracks.CellSelected(1, 2) = True
      End If
   End If
txtCDDBQuery.Enabled = True
If txtCDDBQuery.Text <> "cddb query 0 -1 0" Then
Timer1.Enabled = True
End If
End Sub

Private Sub cboDrives_Click()
   ShowTracks
End Sub

Private Sub chkCompilation_Click()
   '
   grdTracks.ColumnVisible(3) = (chkCompilation.Value = Checked)
   If (chkCompilation.Value) Then
      txtArtist.Tag = txtArtist.Text
      txtArtist.Text = "Various Artists"
   Else
      txtArtist.Text = txtArtist.Tag
   End If
   '
End Sub

Private Sub Configure()
   Dim fO As New frmOptions
   fO.CRRip = m_cRip
   fO.Show vbModal, Me
   If (cboDrives.ListCount > 0) Then
      ' Reset active CD ROM
      cboDrives_Click
   End If
End Sub

Private Sub cmdRefresh_Click()
   ShowTracks
End Sub

Private Sub Rip()
Dim lIndex As Long
   
   lIndex = cboDrives.ListIndex + 1
   If (lIndex > 0) Then
      
      Dim cD As cDrive
      Set cD = m_cRip.CDDrive(lIndex)
      If (cD.IsUnitReady) Then
   
         Dim fRip As New frmRipDialog
         fRip.Writer = PluginManagerInstance.WaveWriter
         
         fRip.OutputDir = App.Path
         fRip.RipTOC = m_cToc
         
         Dim lRow As Long
         Dim cT As cTrack
         For lRow = 1 To grdTracks.Rows
            If (grdTracks.CellItemData(lRow, 1)) Then
               
               Set cT = New cTrack
               cT.Init grdTracks.CellText(lRow, 4), _
                  grdTracks.CellText(lRow, 2), txtAlbum.Text, txtArtist.Text, _
                  IIf(chkCompilation.Value = vbChecked, grdTracks.CellText(lRow, 3), txtArtist.Text), _
                  grdTracks.CellText(lRow, 4), txtYear.Text, ""
               
               fRip.AddTrack cT
               
            End If
         Next
         fRip.Icon = Me.Icon
         
         EnableControls False
                  
         fRip.Show , Me
         Me.Refresh
         DoEvents
                  
         On Error Resume Next
         fRip.RipSelected
         
         Dim lErr As Long, sErr As String
         lErr = Err.Number
         sErr = Err.Description
         If Not (Err.Number = 0) Then
            Unload fRip
            On Error GoTo 0
            MsgBox "An error occurred during ripping: " & sErr, vbExclamation
         End If
         
         EnableControls True

      End If
      
   End If
End Sub



Private Sub Form_Load()
   
   ConfigureGrid
   
   Me.Show
   Me.Refresh
   
    
   
   ' Load default plugins
   With PluginManagerInstance
      .AddPlugin "vbalWaveDataWriter6.cWavFileDataWriter"
      .AddPlugin "vbalMp3DataWriter6.cMp3FileDataWriter"
      .SelectedPluginIndex = 2
   End With
   
   ShowDrives
   Set m_cFreeDb = New cFreeDB
End Sub

Private Sub Form_Resize()
   pnlMain.Height = Me.ScaleHeight
End Sub

Private Sub grdTracks_CancelEdit()
   '
   txtEdit.Visible = False
   '
End Sub

Private Sub grdTracks_PreCancelEdit(ByVal lRow As Long, ByVal lCol As Long, newValue As Variant, bStayInEditMode As Boolean)
   '
   grdTracks.CellText(lRow, lCol) = txtEdit.Text
   '
End Sub

Private Sub grdTracks_RequestEdit(ByVal lRow As Long, ByVal lCol As Long, ByVal iKeyAscii As Integer, bCancel As Boolean)
   '
   If (lCol = 2) Or (lCol = 3) Then
      
      Dim lLeft As Long
      Dim lTop As Long
      Dim lWidth As Long
      Dim lHeight As Long
      
      grdTracks.CellBoundary lRow, lCol, lLeft, lTop, lWidth, lHeight
      lLeft = lLeft + Screen.TwipsPerPixelX + grdTracks.Left
      lTop = lTop + 2 * Screen.TwipsPerPixelY + grdTracks.Top
      txtEdit.Move lLeft, lTop, lWidth, lHeight
      If (iKeyAscii >= 32) Then
         txtEdit.Text = Chr(iKeyAscii)
         txtEdit.SelStart = 1
      Else
         txtEdit.Text = grdTracks.CellText(lRow, lCol)
         txtEdit.SelStart = 0
         txtEdit.SelLength = Len(txtEdit.Text)
      End If
      txtEdit.Visible = True
      txtEdit.ZOrder
      txtEdit.SetFocus
      bCancel = False
      
   Else
      If (lCol = 1) Then
         SelectRow lRow, Not (grdTracks.CellItemData(lRow, 1))
      End If
      
      bCancel = True
            
   End If
   '
End Sub

Private Sub SelectRow(ByVal lRow As Long, ByVal bCheck As Boolean)
Dim lCol As Long
Dim oColor As OLE_COLOR
   oColor = BlendColor(vbHighlight, vbWindowBackground, 32)
   grdTracks.Redraw = False
   grdTracks.CellItemData(lRow, 1) = bCheck
   grdTracks.CellText(lRow, 1) = IIf(bCheck, "b", "")
   For lCol = 1 To grdTracks.Columns
      If (bCheck) Then
         grdTracks.CellBackColor(lRow, lCol) = oColor
      Else
         grdTracks.CellBackColor(lRow, lCol) = -1
      End If
   Next lCol
   grdTracks.Redraw = True
End Sub
Private Sub lblCDDBQuery_Click()
m_cFreeDb.Command = txtCDDBQuery.Text
m_cFreeDb.Start
End Sub
Private Sub mnuEdit_Click(index As Integer)
Dim lRow As Long
   Select Case index
   Case 0
      For lRow = 1 To grdTracks.Rows
         SelectRow lRow, True
      Next lRow
   Case 1
      For lRow = 1 To grdTracks.Rows
         SelectRow lRow, Not (grdTracks.CellItemData(lRow, 1))
      Next lRow
   End Select
End Sub

Private Sub mnuHelp_Click(index As Integer)
   Select Case index
   Case 0
      Dim fA As New frmAbout
      Dim sAck As String
      sAck = "This sample uses components from CDEx, Copyright © 1999 Albert L. Faber and Monty (xiphmont@mit.edu).  "
      sAck = sAck & "CDEx is released under the GNU General Public License and the source code is available from "
      sAck = sAck & "http://cdexos.sourceforge.net."
      fA.Acknowledgements = sAck
      fA.Show vbModal, Me
   End Select
End Sub

Private Sub mnuFile_Click(index As Integer)
   Select Case index
   Case 0
      Rip
   Case 2
      Configure
   Case 4
      Unload Me
   End Select
End Sub

'Private Sub lvwTracks_Click()
'Dim itm As ListItem
'Dim bSelection As Boolean
'
'   For Each itm In lvwTracks.ListItems
'      If (itm.Selected) Then
'         bSelection = True
'      End If
'   Next
'   cmdRip.Enabled = bSelection
'
'End Sub

Private Sub pnlMain_Resize()
   '
   On Error Resume Next
   Dim lHeight As Long
   lHeight = pnlMain.ScaleHeight - grdTracks.Top - txtCDDBQuery.Height - 4 * Screen.TwipsPerPixelY
   grdTracks.Move grdTracks.Left, grdTracks.Top, pnlMain.ScaleWidth - grdTracks.Left * 2, lHeight
   cboDrives.Width = pnlMain.ScaleWidth - cboDrives.Left - cmdRefresh.Width - 2 * Screen.TwipsPerPixelX - grdTracks.Left
   cmdRefresh.Left = cboDrives.Left + cboDrives.Width + 2 * Screen.TwipsPerPixelX
   lblCDDBQuery.Top = grdTracks.Top + grdTracks.Height + 2 * Screen.TwipsPerPixelY
   txtCDDBQuery.Move txtCDDBQuery.Left, lblCDDBQuery.Top, pnlMain.ScaleWidth - txtCDDBQuery.Left - grdTracks.Left
   '
End Sub

Private Sub Timer1_Timer()
Timer1.Enabled = False
lblCDDBQuery_Click
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
   txtEdit.Tag = ""
   If (KeyAscii = 13) Then
      KeyAscii = 0
      grdTracks.EndEdit
      Dim lRow As Long
      Dim lCol As Long
      lRow = grdTracks.SelectedRow
      lCol = grdTracks.SelectedCol
      If (chkCompilation.Value = Checked) Then
         If (lCol = 2) Then
            grdTracks.CellSelected(lRow, 3) = True
            Exit Sub
         End If
      End If
      lRow = lRow + 1
      If (lRow <= grdTracks.Rows) Then
         grdTracks.CellSelected(lRow, 2) = True
      End If
   ElseIf (KeyAscii = 27) Then
      grdTracks.CancelEdit
      KeyAscii = 0
   End If
End Sub
Private Sub txtEdit_LostFocus()
   grdTracks.CancelEdit
End Sub
Private Sub m_cFreeDb_CommandReady()
   Select Case True
   Case InStr(m_cFreeDb.Command, "cddb query") > 0
      Dim cQR As cFreeDbQueryResponse
      Set cQR = m_cFreeDb.QueryResponse
      If (cQR.MatchCount = 0) Then
         If (cQR.ReturnCode = QRMoreThanOneMatch) Then
            MsgBox "No matches found for that CD.", vbInformation
         Else
            MsgBox "Failed to retrieve information: " & cQR.ReturnCodeDescription, vbExclamation
         End If
      Else
         Dim iIndex As Long
         If (cQR.MatchCount > 1) Then
            Dim fS As New frmSelectMatch
            fS.QueryResponse = cQR
            fS.Show vbModal, Me
            iIndex = fS.SelectedMatch
         Else
            iIndex = 1
         End If
         If (iIndex > 0) Then
            txtArtist.Text = cQR.Artist(iIndex)
            txtAlbum.Text = TrimNull(cQR.Title(iIndex))
            'lstDetails.AddItem "Category:" & cQR.Category(iIndex)
            'lstDetails.AddItem ""
            m_cFreeDb.Command = "cddb read " & cQR.Category(iIndex) & " " & cQR.DiscID(iIndex)
            m_cFreeDb.Start
         End If
      End If
      
   Case InStr(m_cFreeDb.Command, "cddb read") > 0
      showCDDetails m_cFreeDb.ReadResponse
   End Select

End Sub
Private Sub showCDDetails( _
      cRR As cFreeDbReadResponse _
   )
With cRR
      Dim i As Long
      For i = 1 To .TrackCount
         grdTracks.CellText(i, 2) = .Title(i)
      Next i
   End With
grdTracks.AutoWidthColumn (2)
End Sub
Public Function TrimNull(item As String) As String
    Dim pos As Integer
        pos = InStr(item, Chr$(13))
        If pos Then item = Left$(item, pos - 1)
        TrimNull = item
End Function
