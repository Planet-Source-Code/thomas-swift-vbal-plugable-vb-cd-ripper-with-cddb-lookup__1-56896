VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cCDDrives"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum eOperatingSystem
   OS_UNKNOWN = -1
   OS_WIN95 = 0
   OS_WIN98 = 1
   OS_WINNT35 = 2
   OS_WINNT4 = 3
   OS_WIN2K = 4
   OS_WINXP = 5
End Enum

'Version structure
Private Type m_tOSVerSIONINFO
   dwm_tOSVersionInfoSize As Long
   dwMajorVersion As Long
   dwMinorVersion As Long
   dwBuildNumber As Long
   dwPlatformId As Long
   szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type
Private Declare Function GetVersionEx Lib "kernel32" _
    Alias "GetVersionExA" _
    (lpVersionInformation As m_tOSVerSIONINFO) As Long

'dwPlatformId defines
Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

' Load Library:
Private Declare Function LoadLibraryEx Lib "kernel32" _
    Alias "LoadLibraryExA" _
    (ByVal lpLibFileName As String, ByVal hFile As Long, _
     ByVal dwFlags As Long) As Long

Private Declare Function FreeLibrary Lib "kernel32" _
    (ByVal hLibModule As Long) As Long

Private Declare Function LoadLibrary Lib "kernel32" _
    Alias "LoadLibraryA" _
    (ByVal lpLibFileName As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" _
    (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

Private m_eOperatingSystem As eOperatingSystem
Private m_tOSVer As m_tOSVerSIONINFO
Private m_bAspiInstalled As Boolean
Private m_iCount As Long
Private Type tCDInfo
   HostAdaptor As Long
   id As Long
   lun As Long
   sDescription As String
   sDriveLetter As String
End Type
Private m_tCDInfo() As tCDInfo

Public Property Get TOC(ByVal index As Long) As cCDToc
   Dim cR As New cCDToc
   cR.fInit m_tCDInfo(index).HostAdaptor, m_tCDInfo(index).id, m_tCDInfo(index).lun
   Set TOC = cR
End Property


Public Sub Refresh()
   
   ' Clear info:
   m_iCount = 0
   Erase m_tCDInfo

   ' get the details:
   If (m_bAspiInstalled) Then
      Dim lAdaptorCount As Long
      Dim lAdaptor As Long
      Dim lDevice As Long
      Dim lMaxDevice As Long
      Dim lLun As Long
      Dim lR As Long
      Dim Inquiry As SRB_HAInquiry
      Dim DevType As SRB_GetDevType
      
      lAdaptorCount = AspiGetNumAdaptors()
      
      If lAdaptorCount > 0 Then
         'scan adapters for CDROMS
         For lAdaptor = 0 To lAdaptorCount - 1
    
            'set up inquiry data
            Inquiry.SRB_Cmd = SC_HA_INQUIRY
            Inquiry.SRB_HaID = lAdaptor
            Inquiry.SRB_Flags = 0
            Inquiry.SRB_Hdr_Rsvd = 0
   
            lR = SendASPI32InquiryEx(Inquiry)
            If (Inquiry.SRB_Status = SS_COMP) Then
            
               ' determine the max target number for the adapter
               ' from offset 3.   If its out of range then we
               ' assume 8 devices:
               lMaxDevice = Inquiry.HA_Unique(3)
               If (lMaxDevice <= 0) Or (lMaxDevice > 8) Then
                  lMaxDevice = 8
               End If

               'scan for CDROM's
               For lDevice = 0 To lMaxDevice - 1
               
                  ' try all 8 values for LUN:
                  For lLun = 0 To 7
               
                     'scan dev types
                     ReDim b(0 To LenB(DevType) - 1) As Byte
                     CopyMemory DevType, b(0), LenB(DevType)
                     
                     DevType.SRB_Cmd = SC_GET_DEV_TYPE
                     DevType.SRB_HaID = lAdaptor
                     DevType.SRB_Target = lDevice
                     DevType.SRB_Lun = lLun
    
                     lR = SendASPI32DevTypeEx(DevType)
                     If (DevType.SRB_Status = SS_COMP) Then
        
                        If DevType.DEV_DeviceType = DTYPE_CDROM Then
                           m_iCount = m_iCount + 1
                           ReDim Preserve m_tCDInfo(1 To m_iCount) As tCDInfo
                           m_tCDInfo(m_iCount).HostAdaptor = lAdaptor
                           m_tCDInfo(m_iCount).id = lDevice
                           m_tCDInfo(m_iCount).lun = lLun
                           fillCDInfo m_tCDInfo(m_iCount)
                        
                        End If
                        
                     End If
                     
                  Next lLun
                     
                  
               Next lDevice
               
            End If
            
        Next lAdaptor
    
      End If
      
   End If
   
End Sub

Private Function AspiGetNumAdaptors() As Long
Dim lR As Long
Dim lStatus As Long
Dim lCount As Long

   lR = GetASPI32SupportInfoEx()
   lStatus = (lR \ &H100&)
   If (lStatus = SS_COMP) Then
      AspiGetNumAdaptors = (lR And &HF&)
   End If

End Function


Private Sub fillCDInfo(tInfo As tCDInfo)
   ' try and find the device name:
   If (m_tOSVer.dwPlatformId = VER_PLATFORM_WIN32_NT) Then
      ' on NT, can look up HKEY_LOCAL_MACHINE\Hardware\DeviceMap\SCSI\
      ' to find the description:
      
   Else
   End If
End Sub

Public Property Get Count() As Long
   Count = m_iCount
End Property

Public Property Get DriveLetter(ByVal nIndex As Long) As String
   DriveLetter = m_tCDInfo(nIndex).sDriveLetter
End Property

Public Property Get Description(ByVal nIndex As Long) As String
   Description = m_tCDInfo(nIndex).sDescription
End Property

Public Property Get HostAdaptor(ByVal nIndex As Long) As Long
   HostAdaptor = m_tCDInfo(nIndex).HostAdaptor
End Property

Public Property Get id(ByVal nIndex As Long) As Long
   id = m_tCDInfo(nIndex).id
End Property

Public Property Get lun(ByVal nIndex As Long) As Long
   lun = m_tCDInfo(nIndex).lun
End Property


Private Sub Class_Initialize()
   
   ' OS Version:
   m_eOperatingSystem = Getm_tOSVersion()
   ' Is ASPI installed?
   m_bAspiInstalled = AspiCheck
    
   Refresh

End Sub
Private Function Getm_tOSVersion() As eOperatingSystem
    
   'get OS version
   m_tOSVer.dwm_tOSVersionInfoSize = Len(m_tOSVer)
   GetVersionEx m_tOSVer
    
   If (m_tOSVer.dwPlatformId = VER_PLATFORM_WIN32_NT) Then
      If (m_tOSVer.dwMajorVersion = 3 And m_tOSVer.dwMinorVersion >= 50) Then
         Getm_tOSVersion = OS_WINNT35
         ' m_tOSVer.dwMajorVersion & "." & m_tOSVer.dwMinorVersion & ", WinNT 3.5"
      ElseIf (m_tOSVer.dwMajorVersion = 4) Then
         Getm_tOSVersion = OS_WINNT4
         ' m_tOSVer.dwMajorVersion & "." & m_tOSVer.dwMinorVersion & ", WinNT 4.0"
      ElseIf (m_tOSVer.dwMajorVersion = 5 And m_tOSVer.dwMinorVersion <= 1) Then
         Getm_tOSVersion = OS_WIN2K
         ' m_tOSVer.dwMajorVersion & "." & m_tOSVer.dwMinorVersion & ", Win2000"
      Else
         Getm_tOSVersion = OS_WINXP
         ' m_tOSVer.dwMajorVersion & "." & m_tOSVer.dwMinorVersion & ", WinXP"
      End If
   ElseIf (m_tOSVer.dwPlatformId = VER_PLATFORM_WIN32_WINDOWS) Then
      If (m_tOSVer.dwMinorVersion = 0) Then
         Getm_tOSVersion = OS_WIN95
         ' m_tOSVer.dwMajorVersion & "." & m_tOSVer.dwMinorVersion & ", Win95"
      Else
         Getm_tOSVersion = OS_WIN98
      ' m_tOSVer.dwMajorVersion & "." & m_tOSVer.dwMinorVersion & ", Win98"
      End If
   Else
      Getm_tOSVersion = OS_UNKNOWN
   End If
    
End Function

Private Function AspiCheck() As Boolean
   Dim hLoad As Long
    
   'load the error messages to parse...
   hLoad = LoadLibrary("WNASPI32.DLL")

   'check for ASPI driver
   If Not (GetProcAddress(hLoad, "GetASPI32SupportInfo") = 0) And _
      Not (GetProcAddress(hLoad, "SendASPI32Command") = 0) Then
      AspiCheck = True
   End If

   If Not (hLoad = 0) Then
      FreeLibrary hLoad
   End If

End Function
