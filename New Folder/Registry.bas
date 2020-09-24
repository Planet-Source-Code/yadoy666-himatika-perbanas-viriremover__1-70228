Attribute VB_Name = "Registry"
'mdlRegistry

'Registry API
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long         ' Note that if you declare the lpData parameter as String, you must pass it By Value.
Public Declare Function RegSetValue Lib "advapi32.dll" Alias "RegSetValueA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal dwType As Long, ByVal lpData As String, ByVal cbData As Long) As Long

Public Const REG_DWORD = 4

Enum REG
    HKEY_CURRENT_USER = &H80000001
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
End Enum

Enum TypeStringValue
    REG_SZ = 1
    REG_EXPAND_SZ = 2
    REG_MULTI_SZ = 7
End Enum

'Create or Set Dword Value Registry
 Public Function CreateDwordValue(hKey As REG, Subkey As String, strValueName As String, dwordData As Long) As Long
 
    On Error Resume Next
    Dim Ret As Long
    
    RegCreateKey hKey, Subkey, Ret
    CreateDwordValue = RegSetValueEx(Ret, strValueName, 0, REG_DWORD, dwordData, 4)
    RegCloseKey Ret
    
End Function

'Create or Set String Value Registry
Public Function CreateStringValue(hKey As REG, Subkey As String, RTypeStringValue As TypeStringValue, strValueName As String, strData As String) As Long
    
    On Error Resume Next
    Dim Ret As Long
    
    RegCreateKey hKey, Subkey, Ret
    CreateStringValue = RegSetValueEx(Ret, strValueName, 0, RTypeStringValue, ByVal strData, Len(strData))
    RegCloseKey Ret
    
End Function

'=======================================================================================
' Ini merupakan fungsi untuk menghapus Value Registry
'=======================================================================================
Public Function DeleteValue(hKey As REG, Subkey As String, lpValName As String) As Long
Dim Ret As Long

    On Error Resume Next
    RegOpenKey hKey, Subkey, Ret
    DeleteValue = RegDeleteValue(Ret, lpValName)
    RegCloseKey Ret
    
End Function


Public Sub CleanReg()

    On Error Resume Next
    'mengaktifkan kembali Regsitry Tools
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System", "DisableRegistryTools", 0
    'mengaktifkan Task manager
        CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System\", "DisableTaskMgr", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System\", "DisableTaskMgr", 0
    'mengaktifkan kembali CMD
    CreateDwordValue HKEY_CURRENT_USER, "Software\Policies\Microsoft\Windows\System", "DisableCMD", 0
    'memunculkan kembali Folder Options
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"
    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFolderOptions"

    DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", "Shell"
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", REG_SZ, "Shell", "Explorer.exe"
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", REG_SZ, "System", ""
    CreateStringValue HKEY_CURRENT_USER, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", REG_SZ, "Shell", "Explorer.exe"

    
    
    'memperbaiki registry yang dirusak
    CreateStringValue HKEY_CLASSES_ROOT, "exefile\shell\open\command", REG_SZ, "", Chr$(34) & "%1" & Chr$(34) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "lnkfile\shell\open\command", REG_SZ, "", Chr$(34) & "%1" & Chr$(34) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "piffile\shell\open\command", REG_SZ, "", Chr$(34) & "%1" & Chr$(34) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "batfile\shell\open\command", REG_SZ, "", Chr$(34) & "%1" & Chr$(34) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "comfile\shell\open\command", REG_SZ, "", Chr$(34) & "%1" & Chr$(34) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "cmdfile\shell\Open\Command", REG_SZ, "", Chr$(34) & "%1" & Chr$(34) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "scrfile\shell\Open\Command", REG_SZ, "", Chr$(34) & "%1" & Chr$(34) & " %*"
    CreateStringValue HKEY_CLASSES_ROOT, "regfile\shell\Open\Command", REG_SZ, "", "regedit.exe %1"

    CreateStringValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SafeBoot\", REG_SZ, "AlternateShell", "cmd.exe"
    CreateStringValue HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet001\Control\SafeBoot\", REG_SZ, "AlternateShell", "cmd.exe"
    CreateStringValue HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet002\Control\SafeBoot\", REG_SZ, "AlternateShell", "cmd.exe"
    CreateStringValue HKEY_LOCAL_MACHINE, "SYSTEM\ControlSet003\Control\SafeBoot\", REG_SZ, "AlternateShell", "cmd.exe"
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon", REG_SZ, "Userinit", GetSystemPath & "userinit.exe"
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", REG_SZ, "Debugger", Chr(&H22) & Left(GetWindowsPath, 3) & "Program Files\Microsoft Visual Studio\Common\MSDev98\Bin\msdev.exe" & Chr(&H22) & " -p %ld -e %ld"
    CreateStringValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows NT\CurrentVersion\AeDebug", REG_SZ, "Auto", "0"
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "Start_ShowControlPanel ", 1
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\WinOldApp\", "Disabled", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\WinOldApp\", "Disabled", 0
    CreateStringValue HKEY_CLASSES_ROOT, "exefile", REG_SZ, "", "Application"
    CreateStringValue HKEY_CLASSES_ROOT, "scrfile", REG_SZ, "", "Screen Saver"
    CreateStringValue HKEY_CLASSES_ROOT, "comfile", REG_SZ, "", "Application"
 
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableConfig", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows NT\SystemRestore", "DisableSR", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "LimitSystemRestoreCheckpointing", 0
    CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Policies\Microsoft\Windows\Installer", "DisableMSI", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoClose", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDesktop", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDriveTypeAutoRun", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoRun", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoLogoff", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoFind", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoTrayContextMenu", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoViewContextMenu", 0
    
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispSettingsPage", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispBackgroundPage", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoScrSavPage", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispApprearancePage", 0
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\System", "NoDispCpl", 0
    
    CreateStringValue HKEY_CURRENT_USER, "Control Panel\Desktop\", REG_SZ, "SCRNSAVE.EXE", ""
    'atur registry agar file dengan yang disembunyikan tidak tampil
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "HideFileExt", 1
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "Hidden", 1
    CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "ShowSuperHidden", 1
    
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrive"
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoDrive"
    DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "DisableRegistryTools"
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "DisableRegistryTools"
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "shutdownwithoutlogon"
    DeleteValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "undockwithoutlogon"
    DoEvents
    
End Sub

Public Sub Reg_default()
On Error Resume Next
'MenuShowDelay
CreateStringValue HKEY_CURRENT_USER, "Control Panel\Desktop", REG_SZ, "MenuShowDelay", "400"
CreateStringValue HKEY_USERS, ".DEFAULT\Control Panel\Desktop", REG_SZ, "MenuShowDelay", "400"
CreateStringValue HKEY_USERS, "S-1-5-18\Control Panel\Desktop", REG_SZ, "MenuShowDelay", "400"
CreateStringValue HKEY_USERS, "S-1-5-19\Control Panel\Desktop", REG_SZ, "MenuShowDelay", "400"
CreateStringValue HKEY_USERS, "S-1-5-20\Control Panel\Desktop", REG_SZ, "MenuShowDelay", "400"
CreateStringValue HKEY_USERS, "S-1-5-21-602162358-261478967-682003330-1003\Control Panel\Desktop", REG_SZ, "MenuShowDelay", "400"

'Unload DLL
CreateDwordValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\", "AlwaysUnloadDLL", 0

'Memory Acces
CreateDwordValue HKEY_LOCAL_MACHINE, "HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\", "DisablePagingExecutive", 0
CreateDwordValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\", "LargeSystemCache", 0

'Autorun CD/Drive
DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoDriveTypeAutoRun"

'Fast Shutdown
DeleteValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\", "WaitToKillServiceTimeout"

'Clear Pagefile
CreateDwordValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Session Manager\Memory Management\", "ClearPageFileAtShutdown", 0

'Control Panel On Start Menu
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "Start_ShowControlPanel ", 1

'Recent Document
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoRecentDocsMenu", 0


'Hidden file
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "Hidden", 0

'OS File
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced\", "ShowSuperHidden", 0

'Menu file
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoFileMenu", 0
CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoFileMenu", 0

'Low Disk
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer\", "NoLowDiskSpaceChecks", 0

'FTP
CreateStringValue HKEY_CURRENT_USER, "Software\Microsoft\FTP", REG_SZ, "Use Web Based FTP", "NO"

'Download
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\", "1803", 1
CreateDwordValue HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\", "1803", 1

'ActiveX
CreateDwordValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Internet Settings\Zones\3\", "1001", 1

End Sub
