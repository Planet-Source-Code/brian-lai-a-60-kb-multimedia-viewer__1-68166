Attribute VB_Name = "ModINI"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal sSectionName As String, ByVal sReturnedString As String, ByVal lSize As Long, ByVal sFileName As String) As Long
Private Declare Function writeprivateprofilestring Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Function ReadINI(Section As String, KeyName As String, filename As String) As String
On Error Resume Next
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName$, "", sRet, Len(sRet), filename))
End Function

Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
On Error Resume Next
    Dim r
    r = writeprivateprofilestring(sSection, sKeyName, sNewString, sFileName)
End Function

Function GetSet(Key As String, Optional Default As String, Optional ForUser As String) As String
On Error Resume Next
Dim Buffer As String
If Len(ForUser) = 0 Then ForUser = UserName
Buffer = ReadINI(ForUser, Key, SettingsFile)
If Len(Buffer) = 0 Then
    If Len(Default) > 0 And Default <> " " Then
        WriteINI ForUser, Key, Default, SettingsFile
    End If
End If
GetSet = ReadINI(ForUser, Key, SettingsFile)
SetAttr SettingsFile, 34 'Hide INI
End Function

Function GetRes(WhichType As String, ControlName As String) As String
    On Error Resume Next
    GetRes = ReadINI(WhichType, ControlName, FindPath(App.Path, "theme.ini"))
End Function

Function SaveSet(Key As String, Value As String, Optional ForUser As String) As String
On Error Resume Next
    If Len(ForUser) = 0 Then ForUser = UserName
    If ReadINI(UserName, "Sandbox", SettingsFile) <> "1" Then 'this stops all ini writes via SaveSet if sandbox is on
        WriteINI ForUser, Key, Value, SettingsFile
    End If
SaveSet = Key
End Function

Public Function SettingsFile() As String
    On Error Resume Next
    SettingsFile = FindPath(App.Path, App.ProductName & ".ini")
End Function

