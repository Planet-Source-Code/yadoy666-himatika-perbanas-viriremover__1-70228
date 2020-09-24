Attribute VB_Name = "Fungsi"
Option Explicit
Public Declare Function lstrcat Lib "Kernel32" Alias "lstrcatA" (ByVal lpString1 As String, _
                                                                 ByVal lpString2 As String) As Long
Public Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BrowseInfo) As Long
Public Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, _
                                                           ByVal lpBuffer As String) As Long
Public Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)

Public Const BIF_RETURNONLYFSDIRS        As Integer = 1
Public Const MAX_PATH                    As Integer = 260


'Loads a MEMORYSTATUS structure with information about the current state of the system’s memory.
Public Declare Sub GlobalMemoryStatus Lib "Kernel32" (lpBuffer As MEMORYSTATUS)

Public Declare Function pBGetFreeSystemResources Lib "rsrc32.dll" _
    Alias "_MyGetFreeSystemResources32@4" (ByVal iResType As Integer) As Integer

Public Declare Function GetVersionEx Lib "Kernel32" _
    Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long

Public Declare Function GetComputerName Lib "kernel32.dll" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long


Const SR = 0
Const GDI = 1
Const USR = 2
Const VER_PLATFORM_WIN32s = 0
Const VER_PLATFORM_WIN32_WINDOWS = 1
Const VER_PLATFORM_WIN32_NT = 2

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion      As Long
    dwMinorVersion      As Long
    dwBuildNumber       As Long
    dwPlatformId        As Long
    szCSDVersion        As String * 128
End Type

Public Type MEMORYSTATUS
    dwLength        As Long ' sizeof(MEMORYSTATUS)
    dwMemoryLoad    As Long ' percent of memory in use
    dwTotalPhys     As Long ' bytes of physical memory
    dwAvailPhys     As Long ' free physical memory bytes
    dwTotalPageFile As Long ' bytes of paging file
    dwAvailPageFile As Long ' free bytes of paging file
    dwTotalVirtual  As Long ' user bytes of address space
    dwAvailVirtual  As Long ' free user bytes
End Type


Public Type BrowseInfo
    lngHwnd                                  As Long
    pIDLRoot                                 As Long
    pszDisplayName                           As Long
    lpszTitle                                As Long
    ulFlags                                  As Long
    lpfnCallback                             As Long
    lParam                                   As Long
    iImage                                   As Long
End Type



Public Function BrowseForFolder(ByVal lngHwnd As Long, _
                                ByVal strPrompt As String) As String

Dim intNull   As Integer
Dim lngIDList As Long
Dim lngResult As Long
Dim strPath   As String
Dim udtBI     As BrowseInfo

    On Error GoTo ehBrowseForFolder
    With udtBI
        .lngHwnd = lngHwnd
        .lpszTitle = lstrcat(strPrompt, "")
        .ulFlags = BIF_RETURNONLYFSDIRS
    End With
    lngIDList = SHBrowseForFolder(udtBI)
    If lngIDList <> 0 Then
        strPath = String$(MAX_PATH, 0)
        lngResult = SHGetPathFromIDList(lngIDList, strPath)
        CoTaskMemFree lngIDList
        intNull = InStr(strPath, vbNullChar)
        If intNull > 0 Then
            strPath = Left$(strPath, intNull - 1)
        End If
    End If
    BrowseForFolder = strPath

Exit Function

ehBrowseForFolder:
    BrowseForFolder = Empty

End Function



