VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScanDirectories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const INVALID_HANDLE_VALUE         As Long = -1
Private Const MAX_PATH                     As Long = 260
Private Const MAXDWORD                     As Long = &HFFFFFFFF
Private Const QS_ALLEVENTS                 As Long = &HBF
Private Type FILETIME
    dwLowDateTime                              As Long
    dwHighDateTime                             As Long
End Type
Private Type SHORTITEMID
    cb                                         As Long
    abID                                       As Integer
End Type
Private Type ITEMIDLIST
    mkid                                       As SHORTITEMID
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes                           As Long
    ftCreationTime                             As FILETIME
    ftLastAccessTime                           As FILETIME
    ftLastWriteTime                            As FILETIME
    nFileSizeHigh                              As Long
    nFileSizeLow                               As Long
    dwReserved0                                As Long
    dwReserved1                                As Long
    cFileName                                  As String * MAX_PATH
    cAlternate                                 As String * 14
End Type
Private m_bCancel                          As Boolean
Private m_bErrorOnFail                     As Boolean
Private m_bFileDeleted                     As Boolean
Private m_bFiltered                        As Boolean
Private m_bFolderDeleted                   As Boolean
Private m_bPaused                          As Boolean
Private m_bScanning                        As Boolean
Private m_bSubDirectories                  As Boolean
Private m_iScanDeep                        As Integer
Private m_iStartDeep                       As Integer
Private m_lAttributes                      As VbFileAttribute
Private m_lFilterUBound                    As Long
Private m_lTotalFiles                      As Long
Private m_lTotalFolders                    As Long
Private m_sCurrentDir                      As String
Private m_sFilters()                       As String
Private m_sStartPath                       As String
Private m_tFindData                        As WIN32_FIND_DATA
Public Event CurrentFile(File As String, Path As String, Delete As Boolean)
Public Event CurrentFolder(Path As String, Cancel As Boolean, Delete As Boolean)
Public Event DoneScanning(TotalFolders As Long, TotalFiles As Long)
#If False Then
Private File, Path, Delete, Cancel, TotalFolders, TotalFiles
#End If
Private Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, _
                                                                                  ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, _
                                                                                ByRef lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function GetQueueStatus Lib "user32.dll" (ByVal fuFlags As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32.dll" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, _
                                                                                        ByVal lpszShortPath As String, _
                                                                                        ByVal cchBuffer As Long) As Long
Private Declare Function lstrlen Lib "kernel32.dll" Alias "lstrlenA" (ByVal lpString As String) As Long
Private Declare Function RemoveDirectory Lib "kernel32.dll" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long

Public Property Get Attributes() As VbFileAttribute

    Attributes = m_lAttributes

End Property

Public Property Let Attributes(Value As VbFileAttribute)

    m_lAttributes = Value

End Property

Public Sub BeginScanning(Optional StartPath, _
                         Optional Filter)

    If m_bPaused Then
        ResumeScanning
    Else
        If Not m_bScanning Then
            m_lTotalFiles = 0
            m_lTotalFolders = 0
            If Not IsMissing(Filter) Then
                Me.Filter = Filter
            End If
            If Not IsMissing(StartPath) Then
                Me.StartPath = StartPath
            End If
            If KeyExists(Me.StartPath) Then
                m_bCancel = False
                StartScanning m_sStartPath
            ElseIf (m_bErrorOnFail) Then
                Err.Raise 76
            End If
            RaiseEvent DoneScanning(m_lTotalFolders, m_lTotalFiles)
        End If
    End If

End Sub

Public Sub CancelScanning()

    m_bCancel = True
    ResumeScanning

End Sub

Private Sub Class_Initialize()

    m_lAttributes = vbNormal And vbHidden And vbReadOnly And vbArchive And vbSystem
    Filter = "*"
    m_sStartPath = CurDir$()
    m_bSubDirectories = True

End Sub

Private Sub Class_Terminate()

    If (m_bScanning Or m_bPaused) Then
        CancelScanning
    End If

End Sub

Public Property Get CurrentFileAttribute() As VbFileAttribute

    CurrentFileAttribute = m_tFindData.dwFileAttributes

End Property

Public Property Get CurrentFileScanned() As String

    CurrentFileScanned = m_tFindData.cFileName

End Property

Public Property Get CurrentFileSize() As Long

    CurrentFileSize = (m_tFindData.nFileSizeHigh * MAXDWORD) + m_tFindData.nFileSizeLow

End Property

Public Property Get CurrentFolderScanned() As String

    CurrentFolderScanned = m_sCurrentDir

End Property

Public Property Get ErrorOnFail() As Boolean

    ErrorOnFail = m_bErrorOnFail

End Property

Public Property Let ErrorOnFail(ByVal Value As Boolean)

    m_bErrorOnFail = Value

End Property

Public Property Get FileDeleted() As Boolean

    FileDeleted = m_bFileDeleted

End Property

Public Property Get Filter() As String

    Filter = Join(m_sFilters(), "|") ' Return as one string

End Property

Public Property Let Filter(Value As String)

    If Len(Value) = 0 Then
        Value = "*"
    End If
    m_bFiltered = Not ((StrComp(Value, "*") = 0))
    m_sFilters = Split(UCase$(Value), "|")
    m_lFilterUBound = UBound(m_sFilters)

End Property

Public Property Get FolderDeleted() As Boolean

    FolderDeleted = m_bFolderDeleted

End Property

Private Function GetShortPath(LongPath As String) As String

Dim p_Buffer As String * MAX_PATH
Dim p_Ret    As Long

    p_Ret = GetShortPathName(LongPath, p_Buffer, MAX_PATH)
    GetShortPath = Left$(p_Buffer, p_Ret)

End Function

Private Function KeyDeep(ByVal Key As String) As Integer

Dim i As Integer

    If Len(Key) Then
        KeyDeep = 1
    End If
    i = 1
    Do While (i)
        i = InStr(i + 1, Key, "\")
        If i Then
            KeyDeep = KeyDeep + 1
        End If
    Loop

End Function

Private Function KeyExists(ByVal Path As String) As Boolean

    If (Not Path Like "*\") Then
        Path = Path & "\"
    End If
    Path = Path & "nul"
    If Len(Dir(Path)) Then
        KeyExists = True
    End If

End Function

Public Property Get Paused() As Boolean

    Paused = m_bPaused

End Property

Public Sub PauseScanning()

    If m_bScanning Then
        m_bPaused = True
    End If

End Sub

Public Sub ResumeScanning()

    If m_bPaused Then
        m_bPaused = False
    End If

End Sub

Public Property Get ScanDeep() As Integer

    ScanDeep = m_iScanDeep

End Property

Public Property Let ScanDeep(ByVal Value As Integer)

    m_iScanDeep = Value

End Property

Private Sub ScanForFiles(Path As String)

Dim i         As Long
Dim j         As Long
Dim p_bDelete As Boolean
Dim p_lFind   As Long
Dim p_lRet    As Long
Dim p_sData   As String

    If Not m_bCancel Then
        On Error Resume Next
        If m_iScanDeep Then
            If (KeyDeep(Path) - m_iStartDeep > m_iScanDeep) Then
                Exit Sub
            End If
        End If
        m_sCurrentDir = Path
        m_lTotalFolders = m_lTotalFolders + 1
        RaiseEvent CurrentFolder(Path, m_bCancel, p_bDelete)
        Path = m_sCurrentDir
        If p_bDelete Then
            m_bFolderDeleted = (RemoveDirectory(Path) <> 0)
            If m_bFolderDeleted Then
                Exit Sub
            End If
        End If
        If (Len(Path) - 2 > MAX_PATH) Then
            p_lFind = FindFirstFile(GetShortPath(Path & "\*"), m_tFindData)
        Else
            p_lFind = FindFirstFile(Path & "\*", m_tFindData)
        End If
        If p_lFind <> INVALID_HANDLE_VALUE Then
            p_lRet = p_lFind
            Do While (p_lRet)
                If GetQueueStatus(QS_ALLEVENTS) Then
                    DoEvents
                End If
                If m_bPaused Then
                    m_bScanning = False
                    Do While (m_bPaused)
                        DoEvents
                    Loop
                    m_bScanning = True
                End If
                If m_bCancel Then
                    FindClose p_lFind
                    Exit Sub
                End If
                p_sData = m_tFindData.cFileName
                p_sData = Left$(p_sData, lstrlen(p_sData))
                If j <> 3 Then
                    j = j + 1
                End If
                If j = 3 Then
                    If (m_tFindData.dwFileAttributes And vbDirectory) = vbDirectory Then
                        If m_bSubDirectories Then
                            ScanForFiles Path & "\" & p_sData
                        End If
                    ElseIf (m_tFindData.dwFileAttributes Or m_lAttributes) Then
                        If Not m_bFiltered Then
                            m_lTotalFiles = m_lTotalFiles + 1
                            RaiseEvent CurrentFile(p_sData, Path, p_bDelete)
                        Else
                            For i = 0 To m_lFilterUBound
                                If (UCase$(p_sData) Like m_sFilters(i)) Then
                                    m_lTotalFiles = m_lTotalFiles + 1
                                    RaiseEvent CurrentFile(p_sData, Path, p_bDelete)
                                End If
                            Next i
                        End If
                        If p_bDelete Then
                            p_bDelete = False
                            m_bFileDeleted = (DeleteFile(Path & "\" & p_sData) <> 0)
                        End If
                    End If
                End If
                p_lRet = FindNextFile(p_lFind, m_tFindData)
            Loop
            FindClose p_lFind
        End If
        On Error GoTo 0
    End If

End Sub

Public Property Get Scanning() As Boolean

    Scanning = m_bScanning

End Property

Public Property Get StartPath() As String

    StartPath = m_sStartPath

End Property

Public Property Let StartPath(ByVal Value As String)

    m_sStartPath = Value

End Property

Private Sub StartScanning(Path As String)

Dim p_lFind As Long

    If (m_lAttributes And vbDirectory) = vbDirectory Then
        m_lAttributes = m_lAttributes - vbDirectory ' Must be a file
    End If
    m_iStartDeep = KeyDeep(Path)
    m_bScanning = True
    If (Len(Path) - 2 > MAX_PATH) Then
        p_lFind = FindFirstFile(GetShortPath(Path & "\*"), m_tFindData)
    Else
        p_lFind = FindFirstFile(Path & "\*", m_tFindData)
    End If
    If p_lFind <> INVALID_HANDLE_VALUE Then
        FindClose p_lFind
    End If
    ScanForFiles Path
    m_bScanning = False

End Sub

Public Property Get SubDirectories() As Boolean

    SubDirectories = m_bSubDirectories

End Property

Public Property Let SubDirectories(ByVal Value As Boolean)

    m_bSubDirectories = Value

End Property

Public Property Get TotalFiles() As Long

    TotalFiles = m_lTotalFiles

End Property

Public Property Get TotalFolders() As Long

    TotalFolders = m_lTotalFolders

End Property






