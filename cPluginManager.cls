VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cPluginManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_colPluginProgIds As New Collection
Private m_colPlugins As New Collection
Private m_iSelPlugin As Long
Private m_sFile As String

Public Property Get File() As String
   File = m_sFile
End Property
Public Property Let File(ByVal sFile As String)
   m_sFile = sFile
   Load
End Property

Public Function AddPlugin(ByVal sProgId As String) As Boolean
   '
   On Error Resume Next
   Dim sChk As String
   sChk = m_colPluginProgIds(sProgId)
   If (Err.Number = 0) Then
      MsgBox "This plugin has already been loaded.", vbExclamation
   End If
   Err.Clear
   
   Dim o As Object
   Set o = CreateObject(sProgId)
   If (Err.Number = 0) Then
      ' Find the wave output interface
      Dim iww As IWaveDataWriter
      Set iww = o
      If (Err.Number = 0) Then
         m_colPluginProgIds.Add sProgId, "C" & m_colPluginProgIds.Count + 1
         m_colPlugins.Add o, sProgId
         AddPlugin = True
      Else
         MsgBox "The specified object does not appear to be a valid plugin [does not implement the IWaveDataWriter interface]", vbExclamation
      End If
   Else
      MsgBox "Could not create an instance of the object '" & sProgId & "'" & vbCrLf & vbCrLf & _
         "The object may not be registered", vbInformation
   End If
   '
End Function

Public Property Get PluginCount() As Long
   PluginCount = m_colPlugins.Count
End Property

Public Property Get PluginProgId(ByVal nIndex As Long) As String
   PluginProgId = m_colPluginProgIds(nIndex)
End Property

Public Property Get PluginInformationByIndex(ByVal nIndex As Long) As IPluginInformation
   
   Dim ii As IPluginInformation
   On Error Resume Next
   Set ii = m_colPlugins.Item(nIndex)
   If (Err.Number = 0) Then
      Set PluginInformationByIndex = ii
   Else
      ' Not supported interface
      Set PluginInformation = Nothing
   End If
   
End Property

Public Property Get SelectedPluginIndex() As Long
   SelectedPluginIndex = m_iSelPlugin
End Property

Public Property Let SelectedPluginIndex(ByVal lIndex As Long)
   m_iSelPlugin = lIndex
End Property

Public Property Get WaveWriter() As IWaveDataWriter
   '
   Set WaveWriter = m_colPlugins.Item(m_iSelPlugin)
   '
End Property

Public Property Get PluginInformation() As IPluginInformation
   '
   Set PluginInformation = PluginInformationByIndex(m_iSelPlugin)
   '
End Property

Public Sub Save()
   If Len(m_sFile) > 0 Then
      ' save persisted information
   End If
End Sub

Public Sub Load()
   Set m_colPlugins = New Collection
   If Len(m_sFile) > 0 Then
      ' read persisted information
   End If
End Sub
