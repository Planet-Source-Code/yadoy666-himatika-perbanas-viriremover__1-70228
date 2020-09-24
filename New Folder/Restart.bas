Attribute VB_Name = "Restart"
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LUID) As Long
Public Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Public Declare Function GetCurrentProcess Lib "Kernel32" () As Long
Public Declare Function GetVersionEx Lib "Kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Const EWX_FORCE = 4
Public Const EWX_REBOOT = 2
Public Const EWX_SHUTDOWN = 1
Public Const VER_PLATFORM_WIN32_NT = 2
Public Const ANYSIZE_ARRAY = 1
Public Const TOKEN_ADJUST_PRIVILEGES = &H20
Public Const TOKEN_QUERY = &H8
Public Const SE_PRIVILEGE_ENABLED = &H2

Public Type LUID
    LowPart As Long
    HighPart As Long
End Type

Public Type LUID_AND_ATTRIBUTES
        pLuid As LUID
        Attributes As Long
End Type

Public Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    Privileges(ANYSIZE_ARRAY) As LUID_AND_ATTRIBUTES
End Type

Public Type OSVERSIONINFO
        dwOSVersionInfoSize As Long
        dwMajorVersion As Long
        dwMinorVersion As Long
        dwBuildNumber As Long
        dwPlatformId As Long
        szCSDVersion As String * 128
End Type

'Reboot Windows(Not WinNT)
Public Function Reboot() As Long

    'On Error Resume Next
    LogOff = ExitWindowsEx(EWX_FORCE Or EWX_REBOOT, 0)
    
End Function

'Shutdown Windows(Not WinNT)
Public Function Shutdown() As Long

    'On Error Resume Next
    LogOff = ExitWindowsEx(EWX_FORCE Or EWX_SHUTDOWN, 0)
    
End Function

'Detection WinNT
Public Function IsWinNT() As Boolean
    
    'On Error Resume Next
    Dim myOS As OSVERSIONINFO

    myOS.dwOSVersionInfoSize = Len(myOS)
    GetVersionEx myOS
    IsWinNT = (myOS.dwPlatformId = VER_PLATFORM_WIN32_NT)
    
End Function

'For Get Privileges from Win NT
Public Sub EnableShutDown()
    
    'On Error Resume Next
    Dim hProc As Long
    Dim hToken As Long
    Dim mLUID As LUID
    Dim mPriv As TOKEN_PRIVILEGES
    Dim mNewPriv As TOKEN_PRIVILEGES
    
    hProc = GetCurrentProcess()
    OpenProcessToken hProc, TOKEN_ADJUST_PRIVILEGES + TOKEN_QUERY, hToken
    LookupPrivilegeValue "", "SeShutdownPrivilege", mLUID
    mPriv.PrivilegeCount = 1
    mPriv.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
    mPriv.Privileges(0).pLuid = mLUID
    'Setting Privileges windows NT
    AdjustTokenPrivileges hToken, False, mPriv, 4 + (12 * mPriv.PrivilegeCount), mNewPriv, 4 + (12 * mNewPriv.PrivilegeCount)

End Sub


' Reboot For WinNT
Public Sub RebootNT(Force As Boolean)

    Dim Flags As Long
    Flags = EWX_REBOOT
    If Force Then Flags = Flags + EWX_FORCE
    If IsWinNT Then EnableShutDown
    ExitWindowsEx Flags, 0
    
End Sub

' Shutdown For WinNT
Public Sub ShutdownNT(Force As Boolean)

    Dim Flags As Long
    Flags = EWX_SHUTDOWN
    If Force Then Flags = Flags + EWX_FORCE
    If IsWinNT Then EnableShutDown
    ExitWindowsEx Flags, 0
    
End Sub


