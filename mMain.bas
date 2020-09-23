Attribute VB_Name = "mMain"
Option Explicit

' ------------------------------------------------------------
' Name:   mMain
' Author: Steve McMahon (steve@vbaccelerator.com)
' Date:   2004-05-06
' Description:
' Starting point for the VB MP3 CD Ripper, plus
' utility functions and singleton definitions.
'
' See http://vbaccelerator.com/
' ------------------------------------------------------------

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" ( _
   ByVal OLE_COLOR As Long, _
   ByVal HPALETTE As Long, _
   pccolorref As Long) As Long
Private Const CLR_INVALID = -1

Private m_cPluginManager As New cPluginManager

Public Property Get PluginManagerInstance() As cPluginManager
   Set PluginManagerInstance = m_cPluginManager
End Property

Public Sub Main()
   InitCommonControls
   
   Dim f As New frmVBCDRip
   f.Show
   
End Sub

Public Property Get BlendColor( _
      ByVal oColorFrom As OLE_COLOR, _
      ByVal oColorTo As OLE_COLOR, _
      Optional ByVal Alpha As Long = 128 _
   ) As Long
Dim lCFrom As Long
Dim lCTo As Long
   lCFrom = TranslateColor(oColorFrom)
   lCTo = TranslateColor(oColorTo)

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

' Convert Automation color to Windows color
Private Function TranslateColor(ByVal clr As OLE_COLOR, _
                        Optional hPal As Long = 0) As Long
    If OleTranslateColor(clr, hPal, TranslateColor) Then
        TranslateColor = CLR_INVALID
    End If
End Function


