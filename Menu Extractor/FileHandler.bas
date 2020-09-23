Attribute VB_Name = "FileHandler"
Option Explicit
'API for FileExists function
Private Const INVALID_HANDLE_VALUE = -1
Private Const MAX_PATH = 260
Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As FILETIME
    ftLastAccessTime As FILETIME
    ftLastWriteTime As FILETIME
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternate As String * 14
End Type
Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long
Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
'.bmu association
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Private Declare Sub SHChangeNotify Lib "shell32.dll" (ByVal wEventId As Long, ByVal uFlags As Long, dwItem1 As Any, dwItem2 As Any)
Private Const SHCNE_ASSOCCHANGED = &H8000000
Private Const SHCNF_IDLIST = &H0
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const REG_SZ = 1
Public safesavename As String
Public Sub SaveSettingString(strPath As String, strValue As String, strData As String)
    Dim hCurKey As Long 'write to registry
    Dim lRegResult As Long
    lRegResult = RegCreateKey(HKEY_CLASSES_ROOT, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    lRegResult = RegCloseKey(hCurKey)
End Sub
Public Function GetSettingString(hKey As Long, strPath As String, strValue As String, Optional Default As String) As String
Dim hCurKey As Long
Dim lValueType As Long
Dim strBuffer As String
Dim lDataBufferSize As Long
Dim intZeroPos As Integer
Dim lRegResult As Long
' Read from registry
If Not IsEmpty(Default) Then
  GetSettingString = Default
Else
  GetSettingString = ""
End If
lRegResult = RegOpenKey(hKey, strPath, hCurKey)
lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, lValueType, ByVal 0&, lDataBufferSize)
If lRegResult = 0 Then
  If lValueType = REG_SZ Then
    strBuffer = String(lDataBufferSize, " ")
    lRegResult = RegQueryValueEx(hCurKey, strValue, 0&, 0&, ByVal strBuffer, lDataBufferSize)
     intZeroPos = InStr(strBuffer, Chr$(0))
    If intZeroPos > 0 Then
      GetSettingString = Left$(strBuffer, intZeroPos - 1)
    Else
      GetSettingString = strBuffer
    End If
  End If
Else
End If
lRegResult = RegCloseKey(hCurKey)
End Function
Public Sub AssocBMU() 'Associate to bmu template files and .frm,.ctl right-click menu(Extract Menu)
    If IsAssociated Then Exit Sub
    SaveSettingString ".bmu", "", App.EXEName
    SaveSettingString App.EXEName + "\DefaultIcon", "", IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + App.EXEName + ".exe"
    SaveSettingString App.EXEName + "\shell\open\command", "", IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + App.EXEName + ".exe %1"
    SaveSettingString "VisualBasic.Form\shell\Extract Menu\command", "", IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + App.EXEName + ".exe %1"
    SaveSettingString "VisualBasic.UserControl\shell\Extract Menu\command", "", IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + App.EXEName + ".exe %1"
    'This line forces windows to update icons
    SHChangeNotify SHCNE_ASSOCCHANGED, SHCNF_IDLIST, 0, 0
End Sub
Public Function IsAssociated() As Boolean
    Dim Answer As String
    Answer = GetSettingString(HKEY_CLASSES_ROOT, App.EXEName + "\shell\open\command", "")
    IsAssociated = (Answer = IIf(Right(App.Path, 1) = "\", App.Path, App.Path + "\") + App.EXEName + ".exe %1")
End Function
Public Function FileExists(sSource As String) As Boolean
    If Right(sSource, 2) = ":\" Then
        Dim allDrives As String
        allDrives = Space$(64)
        Call GetLogicalDriveStrings(Len(allDrives), allDrives)
        FileExists = InStr(1, allDrives, Left(sSource, 1), 1) > 0
        Exit Function
    Else
        If Not sSource = "" Then
            Dim WFD As WIN32_FIND_DATA
            Dim hFile As Long
            hFile = FindFirstFile(sSource, WFD)
            FileExists = hFile <> INVALID_HANDLE_VALUE
            Call FindClose(hFile)
        Else
            FileExists = False
        End If
    End If
End Function
Public Sub FileSave(Text As String, filepath As String)
    On Error Resume Next 'Save files
    Dim f As Integer
    f = FreeFile
    Open filepath For Binary As #f
    Put #f, , Text
    Close #f
    Exit Sub
End Sub
Public Function OneGulp(Src As String) As String
    On Error Resume Next 'Open files
    Dim f As Integer, temp As String, fg As Long
    f = FreeFile
    DoEvents
    Open Src For Binary As #f
    temp = String(LOF(f), Chr$(0))
    Get #f, , temp
    Close #f
    OneGulp = temp
End Function
Public Function FileOnly(ByVal filepath As String) As String
    FileOnly = mID$(filepath, InStrRev(filepath, "\") + 1)
End Function
Public Function ChangeExt(ByVal filepath As String, Optional newext As String) As String
    Dim temp As String
    If InStr(1, filepath, ".") = 0 Then
        temp = filepath
    Else
        temp = mID$(filepath, 1, InStrRev(filepath, "."))
        temp = Left(temp, Len(temp) - 1)
    End If
    If newext <> "" Then newext = "." + newext
    ChangeExt = temp + newext
End Function
Public Function SafeSave(Path As String) As String
    Dim mPath As String, mname As String, mTemp As String, mFile As String, mExt As String, m As Integer
    On Error Resume Next 'Gets a unique name
    mPath = mID$(Path, 1, InStrRev(Path, "\"))
    mname = mID$(Path, InStrRev(Path, "\") + 1)
    mFile = Left(mID$(mname, 1, InStrRev(mname, ".")), Len(mID$(mname, 1, InStrRev(mname, "."))) - 1)
    If mFile = "" Then mFile = mname
    mExt = mID$(mname, InStrRev(mname, "."))
    mTemp = ""
    Do
        If Not FileExists(mPath + mFile + mTemp + mExt) Then
            SafeSave = mPath + mFile + mTemp + mExt
            safesavename = mFile + mTemp + mExt
            Exit Do
        End If
        m = m + 1
        mTemp = Right(Str(m), Len(Str(m)) - 1)
    Loop
End Function


