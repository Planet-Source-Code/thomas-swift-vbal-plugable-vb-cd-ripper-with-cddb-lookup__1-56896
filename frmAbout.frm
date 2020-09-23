VERSION 5.00
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "%Title%"
   ClientHeight    =   4320
   ClientLeft      =   4815
   ClientTop       =   3105
   ClientWidth     =   6285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   288
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   419
   ShowInTaskbar   =   0   'False
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   300
      Top             =   3720
   End
   Begin VB.PictureBox picLogo 
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   1035
      Left            =   1365
      MouseIcon       =   "frmAbout.frx":000C
      MousePointer    =   99  'Custom
      Picture         =   "frmAbout.frx":015E
      ScaleHeight     =   1035
      ScaleWidth      =   4800
      TabIndex        =   7
      Tag             =   "http://vbaccelerator.com/"
      ToolTipText     =   "Click to visit vbAccelerator.com - Advanced VB, C# and VB.NET source code."
      Top             =   75
      Width           =   4800
   End
   Begin VB.TextBox txtAcknowledgements 
      ForeColor       =   &H80000015&
      Height          =   1215
      Left            =   1380
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   2280
      Width           =   4815
   End
   Begin VB.PictureBox picNothing 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   1065
      Left            =   60
      ScaleHeight     =   71
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   408
      TabIndex        =   3
      Top             =   60
      Width           =   6120
      Begin VB.Label lblLinkTo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   255
         Left            =   1320
         TabIndex        =   8
         Top             =   810
         Width           =   4815
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   435
      Left            =   4860
      TabIndex        =   0
      Top             =   3660
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   60
      ScaleHeight     =   1215
      ScaleWidth      =   6120
      TabIndex        =   5
      Top             =   60
      Width           =   6120
   End
   Begin VB.Label lblVersion 
      BackStyle       =   0  'Transparent
      Caption         =   "%Version%"
      Height          =   255
      Left            =   1380
      TabIndex        =   6
      Top             =   1980
      Width           =   4755
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000010&
      X1              =   4
      X2              =   412
      Y1              =   240
      Y2              =   240
   End
   Begin VB.Label lblAppName 
      BackStyle       =   0  'Transparent
      Caption         =   "%AppName%"
      Height          =   255
      Left            =   1380
      TabIndex        =   2
      Top             =   1320
      Width           =   4815
   End
   Begin VB.Label lblCopyright 
      BackStyle       =   0  'Transparent
      Caption         =   "%Copyright%"
      Height          =   495
      Left            =   1380
      TabIndex        =   1
      Top             =   3720
      Width           =   3435
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type
Private Type TRIVERTEX
   x As Long
   y As Long
   Red As Integer
   Green As Integer
   Blue As Integer
   Alpha As Integer
End Type
Private Type GRADIENT_RECT
    UpperLeft As Long
    LowerRight As Long
End Type
Private Type GRADIENT_TRIANGLE
    Vertex1 As Long
    Vertex2 As Long
    Vertex3 As Long
End Type
Private Declare Function GradientFill Lib "msimg32" ( _
   ByVal hDC As Long, _
   pVertex As TRIVERTEX, _
   ByVal dwNumVertex As Long, _
   pMesh As GRADIENT_RECT, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long
Private Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" ( _
   ByVal hDC As Long, _
   pVertex As TRIVERTEX, _
   ByVal dwNumVertex As Long, _
   pMesh As GRADIENT_TRIANGLE, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long

Private Const GRADIENT_FILL_RECT_H As Long = 0
Private Const GRADIENT_FILL_RECT_V As Long = 1
Private Const GRADIENT_FILL_TRIANGLE As Long = &H2

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Const BITSPIXEL = 12         '  Number of bits per pixel

Private Type OSVERSIONINFO
   dwVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion(0 To 127) As Byte
End Type
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInfo As OSVERSIONINFO) As Long
Private Const VER_PLATFORM_WIN32_NT = 2

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetObjectAPI Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function CreateDIBSection Lib "gdi32" _
    (ByVal hDC As Long, _
    pBitmapInfo As BITMAPINFO, _
    ByVal un As Long, _
    lplpVoid As Long, _
    ByVal handle As Long, _
    ByVal dw As Long) As Long

Private Type PictDesc
    cbSizeofStruct As Long
    picType As Long
    hImage As Long
    xExt As Long
    yExt As Long
End Type

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Declare Function OleCreatePictureIndirect Lib "OLEPRO32.DLL" ( _
      lpPictDesc As PictDesc, _
      riid As GUID, _
      ByVal fPictureOwnsHandle As Long, _
      ipic As IPicture _
    ) As Long
Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type
Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Type BLENDFUNCTION
  BlendOp As Byte
  BlendFlags As Byte
  SourceConstantAlpha As Byte
  AlphaFormat As Byte
End Type
' BlendOp:
Private Const AC_SRC_OVER = &H0
' AlphaFormat:
Private Const AC_SRC_ALPHA = &H1

Private Declare Function AlphaBlend Lib "msimg32.dll" ( _
  ByVal hdcDest As Long, _
  ByVal nXOriginDest As Long, _
  ByVal nYOriginDest As Long, _
  ByVal nWidthDest As Long, _
  ByVal nHeightDest As Long, _
  ByVal hDcSrc As Long, _
  ByVal nXOriginSrc As Long, _
  ByVal nYOriginSrc As Long, _
  ByVal nWidthSrc As Long, _
  ByVal nHeightSrc As Long, _
  ByVal lBlendFunction As Long _
) As Long

Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hDC As Long, ByVal nBkMode As Long) As Long
    Private Const OPAQUE = 2
    Private Const TRANSPARENT = 1
Private Declare Function DrawTextA Lib "user32" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DrawTextW Lib "user32" (ByVal hDC As Long, ByVal lpStr As Long, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
    Private Const DT_LEFT = &H0&
    Private Const DT_TOP = &H0&
    Private Const DT_CENTER = &H1&
    Private Const DT_RIGHT = &H2&
    Private Const DT_VCENTER = &H4&
    Private Const DT_BOTTOM = &H8&
    Private Const DT_WORDBREAK = &H10&
    Private Const DT_SINGLELINE = &H20&
    Private Const DT_EXPANDTABS = &H40&
    Private Const DT_TABSTOP = &H80&
    Private Const DT_NOCLIP = &H100&
    Private Const DT_EXTERNALLEADING = &H200&
    Private Const DT_CALCRECT = &H400&
    Private Const DT_NOPREFIX = &H800
    Private Const DT_INTERNAL = &H1000&
    Private Const DT_WORD_ELLIPSIS = &H40000


Private m_hDib(0 To 1) As Long
Private m_hBmpOld(0 To 1) As Long
Private m_hDC(0 To 1) As Long
Private m_lPtr(0 To 1) As Long
Private m_tBI(0 To 1) As BITMAPINFO
Private m_lAlpha As Single

Private m_sAppName As String
Private m_sVersion As String
Private m_sCopyright As String
Private m_sCopyrightUrl As String
Private m_sTitle As String
Private m_sAcknowledgements As String
Private m_bShowing As Boolean

Private m_xDir() As Long
Private m_yDir() As Long

Private m_tOSV As OSVERSIONINFO


Private Sub SetAsUrl(ctl As Control, ByVal sUrl As String)
   If (Len(sUrl) > 0) Then
      ctl.Tag = sUrl
      ctl.ForeColor = &H800000
      ctl.FontUnderline = True
      ctl.MouseIcon = picLogo.MouseIcon
      ctl.MousePointer = 99
      If (Len(ctl.ToolTipText) = 0) Then
         ctl.ToolTipText = "Go to " & sUrl
      End If
   Else
      If Len(ctl.Tag) > 0 Then
         If (ctl.ToolTipText = "Go to " & ctl.Tag) Then
            ctl.ToolTipText = ""
         End If
         ctl.Tag = ""
      End If
      ctl.ForeColor = vbWindowText
      ctl.FontUnderline = False
      ctl.MousePointer = vbDefault
   End If
End Sub

Private Sub UrlClick(ctl As Control)
   If Len(ctl.Tag) > 0 Then
      ShellExecute Me.hWnd, "open", ctl.Tag, "", "", SW_SHOWNORMAL
   End If
End Sub

Public Property Get Acknowledgements() As String
   Acknowledgements = m_sAcknowledgements
End Property
Public Property Let Acknowledgements(ByVal sAcknowledgements As String)
   m_sAcknowledgements = sAcknowledgements
   If (m_bShowing) Then
      txtAcknowledgements = m_sAcknowledgements
   End If
End Property
Public Property Get Title() As String
   Title = m_sTitle
End Property
Public Property Let Title(ByVal sTitle As String)
   m_sTitle = sTitle
   If (m_bShowing) Then
      Me.Caption = m_sTitle
   End If
End Property
Public Property Get Copyright() As String
   Copyright = m_sCopyright
End Property
Public Property Let Copyright(ByVal sCopyright As String)
   m_sCopyright = sCopyright
   If (m_bShowing) Then
      lblCopyright.Caption = m_sCopyright
   End If
End Property
Public Property Get CopyrightUrl() As String
   CopyrightUrl = m_sCopyrightUrl
End Property
Public Property Let CopyrightUrl(ByVal sCopyrightUrl As String)
   m_sCopyrightUrl = sCopyrightUrl
   If (m_bShowing) Then
      SetAsUrl lblCopyright, m_sCopyrightUrl
   End If
End Property

Public Property Get Version() As String
   Version = m_sVersion
End Property
Public Property Let Version(ByVal sVersion As String)
   m_sVersion = sVersion
   If (m_bShowing) Then
      lblVersion.Caption = sVersion
   End If
End Property
Public Property Get AppName() As String
   AppName = m_sAppName
End Property
Public Property Let AppName(ByVal sAppName As String)
   m_sAppName = sAppName
   If (m_bShowing) Then
      lblAppName.Caption = m_sAppName
   End If
End Property


Private Sub cmdOK_Click()
   Unload Me
End Sub

Private Sub Form_Load()
   
   m_tOSV.dwVersionInfoSize = Len(m_tOSV)
   GetVersionEx m_tOSV
   
   ' Defaults:
   If Len(m_sAppName) = 0 Then
      m_sAppName = App.FileDescription
   End If
   If Len(m_sTitle) = 0 Then
      m_sTitle = App.Title
   End If
   If Len(m_sCopyright) = 0 Then
      m_sCopyright = App.LegalCopyright
   End If
   If Len(m_sVersion) = 0 Then
      m_sVersion = "Version: " & App.Major & "." & App.Minor & "." & App.Revision
   End If
   If Len(m_sAcknowledgements) = 0 Then
      m_sAcknowledgements = App.Comments
   End If
   
   Me.Caption = m_sTitle
   lblAppName.Caption = m_sAppName
   lblCopyright.Caption = m_sCopyright
   SetAsUrl lblCopyright, m_sCopyrightUrl
   lblVersion.Caption = m_sVersion
   txtAcknowledgements.Text = m_sAcknowledgements
   
   If Me.Icon.handle = 0 Then
      Dim frm As Form
      For Each frm In Forms
         If (frm.BorderStyle = 2) Then
            Me.Icon = frm.Icon
            Exit For
         End If
      Next
   End If
   
   m_bShowing = True
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
   ClearUp 1
   ClearUp 0
End Sub

Private Sub Form_Resize()
   If (m_tOSV.dwMajorVersion > 4) Or _
      (m_tOSV.dwMajorVersion = 4) And (m_tOSV.dwMinorVersion >= 10) Then
      If (GetDeviceCaps(Me.hDC, BITSPIXEL) > 8) Then
         Me.Cls
         Dim tR As RECT
         tR.Left = 0
         tR.Top = 0
         tR.Right = Me.ScaleWidth
         tR.Bottom = Me.ScaleHeight
         GradientFillTri Me.hDC, tR, _
            BlendColor(vb3DHighlight, vbButtonFace), vbButtonFace, _
            vbButtonFace, BlendColor(vbButtonFace, vbButtonShadow, 224)
         Me.Refresh
      End If
   End If
End Sub

Private Sub lblCopyright_Click()
   UrlClick lblCopyright
End Sub

Private Sub lblLinkTo_Click()
   UrlClick lblLinkTo
End Sub

Private Sub picLogo_Click()
   UrlClick picLogo
End Sub

Private Sub GradientFillTri( _
      ByVal lHDC As Long, _
      tR As RECT, _
      ByVal topLeftColor As OLE_COLOR, _
      ByVal topRightColor As OLE_COLOR, _
      ByVal bottomLeftColor As OLE_COLOR, _
      ByVal bottomRightColor As OLE_COLOR _
   )
Dim hBrush As Long
Dim lTopLeftColor As Long
Dim lTopRightColor As Long
Dim lBottomLeftColor As Long
Dim lBottomRightColor As Long
Dim lR As Long
   
   ' Use GradientFill:
   OleTranslateColor topLeftColor, 0, lTopLeftColor
   OleTranslateColor topRightColor, 0, lTopRightColor
   OleTranslateColor bottomLeftColor, 0, lBottomLeftColor
   OleTranslateColor bottomRightColor, 0, lBottomRightColor

   Dim tTV(0 To 3) As TRIVERTEX
   Dim tGR(0 To 1)  As GRADIENT_TRIANGLE
   
   setTriVertexColor tTV(0), lTopLeftColor
   tTV(0).x = tR.Left
   tTV(0).y = tR.Top
   
   setTriVertexColor tTV(1), lTopRightColor
   tTV(1).x = tR.Right
   tTV(1).y = tR.Top
   
   setTriVertexColor tTV(2), lBottomRightColor
   tTV(2).x = tR.Right
   tTV(2).y = tR.Bottom
   
   setTriVertexColor tTV(3), lBottomLeftColor
   tTV(3).x = tR.Left
   tTV(3).y = tR.Bottom
   
   tGR(0).Vertex1 = 0
   tGR(0).Vertex2 = 1
   tGR(0).Vertex3 = 2
    
   tGR(1).Vertex1 = 0
   tGR(1).Vertex2 = 2
   tGR(1).Vertex3 = 3
   
   GradientFillTriangle lHDC, tTV(0), 4, tGR(0), 2, GRADIENT_FILL_TRIANGLE
   
End Sub

Private Sub setTriVertexColor(tTV As TRIVERTEX, lColor As Long)
Dim lRed As Long
Dim lGreen As Long
Dim lBlue As Long
   lRed = (lColor And &HFF&) * &H100&
   lGreen = (lColor And &HFF00&)
   lBlue = (lColor And &HFF0000) \ &H100&
   setTriVertexColorComponent tTV.Red, lRed
   setTriVertexColorComponent tTV.Green, lGreen
   setTriVertexColorComponent tTV.Blue, lBlue
End Sub
Private Sub setTriVertexColorComponent(ByRef iColor As Integer, ByVal lComponent As Long)
   If (lComponent And &H8000&) = &H8000& Then
      iColor = (lComponent And &H7F00&)
      iColor = iColor Or &H8000
   Else
      iColor = lComponent
   End If
End Sub


Private Sub picNothing_Click()
   If (GetDeviceCaps(Me.hDC, BITSPIXEL) > 8) Then
      If Not (tmr.Enabled) Then
         CreateFromHBitmap 0, picLogo.Picture.handle
         Create 1, m_tBI(0).bmiHeader.biWidth, m_tBI(0).bmiHeader.biHeight
         LoadPictureBlt 1, m_hDC(0)
         picLogo.Visible = False
         m_lAlpha = 255
         tmr_Timer
         tmr.Enabled = True
      Else
         picLogo.Visible = True
         tmr.Enabled = False
      End If
   End If
End Sub

Private Sub tmr_Timer()
Dim tSAIn As SAFEARRAY2D
Dim bDibIn() As Byte
Dim tSAOut As SAFEARRAY2D
Dim bDibOut() As Byte
Static lIndexFrom As Long
Static lIndexTo As Long

Dim xEnd As Long, yEnd As Long
Dim x As Long, y As Long, y2 As Long, x2 As Long

   lIndexFrom = lIndexTo
   If (lIndexFrom = 1) Then
      lIndexTo = 0
   Else
      lIndexTo = 1
   End If
   '
   ' Get the bits in the from DIB section:
   With tSAIn
       .cbElements = 1
       .cDims = 2
       .Bounds(0).lLbound = 0
       .Bounds(0).cElements = m_tBI(lIndexFrom).bmiHeader.biHeight
       .Bounds(1).lLbound = 0
       .Bounds(1).cElements = 4 * m_tBI(lIndexFrom).bmiHeader.biWidth
       .pvData = m_lPtr(lIndexFrom)
   End With
   CopyMemory ByVal VarPtrArray(bDibIn()), VarPtr(tSAIn), 4

   With tSAOut
       .cbElements = 1
       .cDims = 2
       .Bounds(0).lLbound = 0
       .Bounds(0).cElements = m_tBI(lIndexTo).bmiHeader.biHeight
       .Bounds(1).lLbound = 0
       .Bounds(1).cElements = 4 * m_tBI(lIndexTo).bmiHeader.biWidth
       .pvData = m_lPtr(lIndexTo)
   End With
   CopyMemory ByVal VarPtrArray(bDibOut()), VarPtr(tSAOut), 4

   xEnd = tSAIn.Bounds(1).cElements - 4
   yEnd = tSAIn.Bounds(0).cElements - 1
   For y = yEnd To 0 Step -1
      For x = 0 To xEnd Step 4
         
         y2 = y + (Rnd * 4) - 2
         If (y2 < 0) Then
            y2 = yEnd + y2
         End If
         If (y2 > yEnd) Then
            y2 = y2 - yEnd - 1
         End If
         
         x2 = x + Int(Rnd * 4) * 4 - 64
         If (x2 < 0) Then
            x2 = xEnd + x2
         End If
         If (x2 > xEnd) Then
            x2 = x2 - xEnd - 4
         End If
         
         bDibOut(x, y) = (14& * bDibIn(x, y)) \ 16& + (1& * bDibIn(x2, y2) \ 16&) + 16&
         bDibOut(x + 1, y) = (14& * bDibIn(x + 1, y)) \ 16& + (1& * bDibIn(x2 + 1, y2) \ 16&) + 16&
         bDibOut(x + 2, y) = (14& * bDibIn(x + 2, y)) \ 16& + (1& * bDibIn(x2 + 2, y2) \ 16&) + 16&
         
      Next x
   Next y
   
   ' Clear the temporary array descriptor
   CopyMemory ByVal VarPtrArray(bDibIn), 0&, 4
   CopyMemory ByVal VarPtrArray(bDibOut), 0&, 4
   
   picNothing.Cls
   Dim lTextAlpha As Long
   Dim tR As RECT
   tR.Left = 87
   tR.Right = 87 + m_tBI(0).bmiHeader.biWidth
   tR.Top = 1
   tR.Bottom = 1 + m_tBI(0).bmiHeader.biHeight
   lTextAlpha = m_lAlpha + 128
   If (lTextAlpha <= 255) Then
      SetBkMode picNothing.hDC, TRANSPARENT
      SetTextColor picNothing.hDC, BlendColor(vbButtonShadow, vb3DHighlight, lTextAlpha)
      DrawText picNothing.hDC, "vbAccelerator is a site providing advanced, free source code for VB, C# and VB.NET programmers.  Specialities are controls, user interface and imaging.  Everything is free and comes complete with full source code.", tR, DT_CENTER Or DT_VCENTER Or DT_WORDBREAK
      If (lTextAlpha < 64) And (lblLinkTo.Tag = "") Then
         lblLinkTo.Caption = "http://vbaccelerator.com/"
         SetAsUrl lblLinkTo, "http://vbaccelerator.com/"
         lblLinkTo.ForeColor = BlendColor(vbButtonShadow, vb3DHighlight, lTextAlpha)
      End If
   End If
   If (m_lAlpha > 0) Then
   
      Dim lBlend As Long
      Dim bf As BLENDFUNCTION
      bf.BlendOp = AC_SRC_OVER
      bf.BlendFlags = 0
      bf.SourceConstantAlpha = m_lAlpha
      bf.AlphaFormat = 0
      CopyMemory lBlend, bf, 4
      AlphaBlend picNothing.hDC, _
         87, 1, m_tBI(0).bmiHeader.biWidth, m_tBI(0).bmiHeader.biHeight, _
         m_hDC(lIndexTo), _
         0, 0, m_tBI(0).bmiHeader.biWidth, m_tBI(0).bmiHeader.biHeight, _
         lBlend
   End If
   
   m_lAlpha = m_lAlpha - 4
   If (m_lAlpha < -128) Then
      tmr.Enabled = False
   End If
   picNothing.Refresh
   '
End Sub

Private Function CreateFromHBitmap( _
      ByVal lIndex As Long, _
      ByVal hBmp As Long _
   )
Dim lHDC As Long
Dim lhWnd As Long
Dim lhDCDesktop As Long
Dim lhBmpOld As Long
Dim tBmp As BITMAP
   GetObjectAPI hBmp, Len(tBmp), tBmp
   If (Create(lIndex, tBmp.bmWidth, tBmp.bmHeight)) Then
      lhWnd = GetDesktopWindow()
      lhDCDesktop = GetDC(lhWnd)
      If (lhDCDesktop <> 0) Then
         lHDC = CreateCompatibleDC(lhDCDesktop)
         ReleaseDC lhWnd, lhDCDesktop ' 2003-07-05: Corrected for GDI leak in Win98
         If (lHDC <> 0) Then
            lhBmpOld = SelectObject(lHDC, hBmp)
            LoadPictureBlt lIndex, lHDC
            SelectObject lHDC, lhBmpOld
            DeleteDC lHDC
         End If
      End If
   End If
   
End Function

Private Function Create( _
        ByVal lIndex As Long, _
        ByVal lWidth As Long, _
        ByVal lHeight As Long _
    ) As Boolean
   ClearUp lIndex
   m_hDC(lIndex) = CreateCompatibleDC(Me.hDC)
   If (m_hDC(lIndex) <> 0) Then
       If (CreateDIB(lIndex, m_hDC(lIndex), lWidth, lHeight, m_hDib(lIndex))) Then
           m_hBmpOld(lIndex) = SelectObject(m_hDC(lIndex), m_hDib(lIndex))
           Create = True
       Else
           DeleteDC m_hDC(lIndex)
           m_hDC(lIndex) = 0
       End If
   End If
End Function

Private Sub LoadPictureBlt( _
        ByVal lIndex As Long, _
        ByVal lHDC As Long, _
        Optional ByVal lSrcLeft As Long = 0, _
        Optional ByVal lSrcTop As Long = 0, _
        Optional ByVal lSrcWidth As Long = -1, _
        Optional ByVal lSrcHeight As Long = -1, _
        Optional ByVal eRop As RasterOpConstants = vbSrcCopy _
    )
    If lSrcWidth < 0 Then lSrcWidth = m_tBI(lIndex).bmiHeader.biWidth
    If lSrcHeight < 0 Then lSrcHeight = m_tBI(lIndex).bmiHeader.biHeight
    BitBlt m_hDC(lIndex), 0, 0, lSrcWidth, lSrcHeight, lHDC, lSrcLeft, lSrcTop, eRop
End Sub

Private Sub ClearUp(ByVal lIndex As Long)
   If (m_hDC(lIndex) <> 0) Then
      If (m_hDib(lIndex) <> 0) Then
         SelectObject m_hDC(lIndex), m_hBmpOld(lIndex)
         DeleteObject m_hDib(lIndex)
      End If
      DeleteObject m_hDC(lIndex)
   End If
   m_hDC(lIndex) = 0: m_hDib(lIndex) = 0: m_hBmpOld(lIndex) = 0: m_lPtr(lIndex) = 0
End Sub

Private Function CreateDIB( _
      ByVal lIndex As Long, _
      ByVal lHDC As Long, _
      ByVal lWidth As Long, _
      ByVal lHeight As Long, _
      ByRef hDib As Long _
   ) As Boolean
   With m_tBI(lIndex).bmiHeader
      .biSize = Len(m_tBI(lIndex).bmiHeader)
      .biWidth = lWidth
      .biHeight = lHeight
      .biPlanes = 1
      .biBitCount = 32
      .biCompression = BI_RGB
      .biSizeImage = lWidth * 4 * lHeight
   End With
   hDib = CreateDIBSection( _
         lHDC, _
         m_tBI(lIndex), _
         DIB_RGB_COLORS, _
         m_lPtr(lIndex), _
         0, 0)
   CreateDIB = (hDib <> 0)
End Function

Private Property Get BlendColor( _
      ByVal oColorFrom As OLE_COLOR, _
      ByVal oColorTo As OLE_COLOR, _
      Optional ByVal Alpha As Long = 128 _
   ) As Long
Dim lCFrom As Long
Dim lCTo As Long
   OleTranslateColor oColorFrom, 0, lCFrom
   OleTranslateColor oColorTo, 0, lCTo
Dim lSrcR As Long
Dim lSrcG As Long
Dim lSrcB As Long
Dim lDstR As Long
Dim lDstG As Long
Dim lDstB As Long
   lSrcR = lCFrom And &HFF
   lSrcG = (lCFrom And &HFF00&) \ &H100&
   lSrcB = (lCFrom And &HFF0000) \ &H10000
   lDstR = lCTo And &HFF
   lDstG = (lCTo And &HFF00&) \ &H100&
   lDstB = (lCTo And &HFF0000) \ &H10000
     
   
   BlendColor = RGB( _
      ((lSrcR * Alpha) / 255) + ((lDstR * (255 - Alpha)) / 255), _
      ((lSrcG * Alpha) / 255) + ((lDstG * (255 - Alpha)) / 255), _
      ((lSrcB * Alpha) / 255) + ((lDstB * (255 - Alpha)) / 255) _
      )
      
End Property

Private Function DrawText(ByVal hDC As Long, ByVal sText As String, rc As RECT, ByVal lFlags As Long)
Dim lPtr As Long
   If (m_tOSV.dwPlatformId = VER_PLATFORM_WIN32_NT) Then
      lPtr = StrPtr(sText)
      If Not (lPtr = 0) Then
         DrawTextW hDC, ByVal lPtr, -1, rc, lFlags
      End If
   Else
      DrawTextA hDC, sText, -1, rc, lFlags
   End If

End Function
