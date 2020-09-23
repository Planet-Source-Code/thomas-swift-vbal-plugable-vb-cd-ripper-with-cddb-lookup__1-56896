VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFreeDB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type QOCINFO
    dwSize As Long
    dwFlags As Long
    dwInSpeed As Long
    dwOutSpeed As Long
End Type
Private Enum NetworkAliveFlags
   NETWORK_ALIVE_LAN = &H1&
   NETWORK_ALIVE_WAN = &H2&
   NETWORK_ALIVE_AOL = &H4&
End Enum
Private Declare Function IsNetworkAlive Lib "Sensapi.dll" ( _
      ByVal lpdwFlags As NetworkAliveFlags _
   ) As Long
Private Declare Function IsDestinationReachableA Lib "Sensapi.dll" ( _
      ByVal lpszDestination As String, _
      lpQOCInfo As Any _
    ) As Long
Private Declare Function FreeLibrary Lib "kernel32" _
    (ByVal hLibModule As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32" _
    Alias "LoadLibraryA" _
    (ByVal lpLibFileName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" _
    (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private WithEvents m_tmr As CTimer
Attribute m_tmr.VB_VarHelpID = -1

Private m_xHttp As ServerXMLHTTP
Public Enum EXmlHttpReadyState
   Uninitialised = 0
   Loading = 1
   Loaded = 2
   Interactive = 3
   Completed = 4
End Enum

Private m_sServer As String
Private m_sUrl As String
Private m_sUserName As String
Private m_sUserHost As String
Private m_sAgentName As String
Private m_sAgentVersion As String
Private m_sCommand As String
Private m_sResponse As String
Private m_bQueryRunning As Boolean

Public Event CommandReady()

Private Property Get SensApiAvailable() As Boolean
Dim hLoad As Long
   hLoad = LoadLibrary("sensapi.dll")
   If Not (hLoad = 0) Then
      If Not (GetProcAddress(hLoad, "IsNetworkAlive") = 0) Then
         SensApiAvailable = True
      End If
      FreeLibrary hLoad
   End If
End Property

Public Property Get CanConnect() As Boolean
   If (SensApiAvailable()) Then ' IE5+ only
      If Not (IsNetworkAlive(NETWORK_ALIVE_WAN) = 0) Then
         Dim sUrl As String
         sUrl = m_sServer & m_sUrl
         If Not (IsDestinationReachableA(sUrl, ByVal 0&) = 0) Then
            CanConnect = True
         End If
      End If
   End If
End Property

Public Property Get Server() As String
   Server = m_sServer
End Property
Public Property Let Server(ByVal Value As String)
   m_sServer = Value
End Property

Public Property Get Url() As String
   Url = m_sUrl
End Property
Public Property Let Url(ByVal Value As String)
   m_sUrl = Value
End Property

Public Property Get UserName() As String
   UserName = m_sUserName
End Property
Public Property Let UserName(ByVal Value As String)
   m_sUserName = Value
End Property

Public Property Get UserHost() As String
   UserHost = m_sUserHost
End Property
Public Property Let UserHost(ByVal Value As String)
   m_sUserHost = Value
End Property

Public Property Get AgentName() As String
   AgentName = m_sAgentName
End Property
Public Property Let AgentName(ByVal Value As String)
   m_sAgentName = Value
End Property

Public Property Get AgentVersion() As String
   AgentVersion = m_sAgentVersion
End Property
Public Property Let AgentVersion(ByVal Value As String)
   m_sAgentVersion = Value
End Property

Public Property Get Command() As String
   Command = m_sCommand
End Property
Public Property Let Command(ByVal Value As String)
   m_sCommand = Value
End Property

Public Property Get Response() As String
   Response = m_sResponse
End Property


Public Sub Start()
   
   If (m_bQueryRunning) Then
      Err.Raise vbObjectError + 4096 + 1, App.EXEName & ".cFreeDB", "Query already in progress"
   Else
      m_sResponse = ""
      Dim sUrl As String
      sUrl = UrlBit & parse(m_sCommand) & HelloBit
      m_xHttp.open "GET", sUrl, True
      m_xHttp.send
      m_bQueryRunning = True
      m_tmr.Interval = 50
   End If
   
End Sub

Public Property Get CommandSubmitted() As Boolean
   CommandSubmitted = m_bQueryRunning
End Property

Public Sub Abort()
   If (m_bQueryRunning) Then
      If (m_bQueryRunning) Then
         m_tmr.Interval = 0
         m_xHttp.Abort
      End If
   End If
End Sub

Private Sub Class_Initialize()
   m_sAgentName = "VBALTrackList"
   m_sAgentVersion = "1.0"
   m_sUserName = "steve"
   m_sUserHost = "vbaccelerator.com"
   m_sServer = "http://freedb.freedb.org/"
   m_sUrl = "~cddb/cddb.cgi"
   Set m_xHttp = New ServerXMLHTTP
   Set m_tmr = New CTimer
End Sub

Private Function UrlBit() As String
Dim sUrlBit As String
   sUrlBit = m_sServer
   sUrlBit = sUrlBit & m_sUrl
   sUrlBit = sUrlBit & "?cmd="
   UrlBit = sUrlBit
End Function

Private Function HelloBit() As String
Dim sHelloBit As String
   sHelloBit = "&hello="
   sHelloBit = sHelloBit & parse(m_sUserName)
   sHelloBit = sHelloBit & "+" & parse(m_sUserHost)
   sHelloBit = sHelloBit & "+" & parse(m_sAgentName)
   sHelloBit = sHelloBit & "+" & parse(m_sAgentVersion)
   sHelloBit = sHelloBit & "&proto=1"
   HelloBit = sHelloBit
End Function

Private Function parse(ByVal sGetUrl) As String
On Error Resume Next
Dim i As Long
Dim b() As Byte
Dim sRet As String
   b = StrConv(sGetUrl, vbFromUnicode)
   For i = LBound(b) To UBound(b)
      Select Case b(i)
      Case 0 To 9
         ' non-printing codes:
         sRet = "%0" & Hex(b)
      Case 10 To 31
         ' non-printing codes:
         sRet = "%" & Hex(b)
      Case 32
         ' space
         sRet = sRet & "+"
      Case 33 To 45, 47
         ' codes which can affect URLs
         sRet = "%" & Hex(b)
      Case 46
         ' full stop
         sRet = sRet & "."
      Case 48 To 57
         ' numbers:
         sRet = sRet & Chr(b(i))
      Case 58 To 64
         ' codes which can affect URLs
         sRet = "%" & Hex(b)
      Case 65 To 90
         ' upper case letters
         sRet = sRet & Chr(b(i))
      Case 91 To 96
         ' codes which affect URLs
         sRet = "%" & Hex(b(i))
      Case 97 To 122
         ' lower case letters
         sRet = sRet & Chr(b(i))
      Case 126 To 255
         ' codes which affect URLs
         sRet = "%" & Hex(b)
      End Select
   Next i
   parse = sRet
End Function

Private Sub Class_Terminate()
   Abort
   Set m_tmr = Nothing
   Set m_xHttp = Nothing
End Sub

Private Sub m_tmr_ThatTime()
   If (m_xHttp.readyState = Completed) Then
      m_tmr.Interval = 0
      m_bQueryRunning = False
      Dim b() As Byte
      b = m_xHttp.responseBody
      m_sResponse = StrConv(b, vbUnicode)
      RaiseEvent CommandReady
   End If
End Sub


Public Property Get QueryResponse() As cFreeDbQueryResponse
   Dim cF As New cFreeDbQueryResponse
   cF.fInit m_sResponse
   Set QueryResponse = cF
End Property

Public Property Get ReadResponse() As cFreeDbReadResponse
   Dim cR As New cFreeDbReadResponse
   cR.fInit m_sResponse
   Set ReadResponse = cR
End Property





