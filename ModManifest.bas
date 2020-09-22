Attribute VB_Name = "ModManifest"
'This Module is in public use with:
'Nothing else but you, ProFile.

Public JustMade As Boolean 'did I just make the manifest and it isn't going to work on this instance...?

Public Type tagInitCommonControlsEx
    lngSize As Long
    lngICC As Long
End Type

Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As tagInitCommonControlsEx) As Boolean
Const ICC_USEREX_CLASSES = &H200

Public Function XPVB() As Boolean 'my manifesting module
    On Error Resume Next
    If Dir(MyManifestFile) <> "" Then GoTo Written
    Dim XPStr As String
    Dim FF As Integer
    XPStr = "<?xml version=""1.0"" encoding=""UTF-8"" standalone=""yes""?>" & vbCrLf & _
            "<assembly xmlns=""urn:schemas-microsoft-com:asm.v1"" manifestVersion=""1.0"">" & vbCrLf & _
            "<assemblyIdentity version=""1.0.0.0"" processorArchitecture=""X86"" name=""" & App.ProductName & """ type=""win32""/>" & vbCrLf & _
            "<description>" & App.ProductName & " manifest file</description>" & vbCrLf & "<dependency>" & vbCrLf & _
            "<dependentAssembly>" & vbCrLf & "<assemblyIdentity type=""win32"" name=""Microsoft.Windows.Common-Controls"" version=""6.0.0.0"" processorArchitecture=""X86"" publicKeyToken=""6595b64144ccf1df"" language=""*""/>" & vbCrLf & _
            "</dependentAssembly>" & vbCrLf & "</dependency>" & vbCrLf & "</assembly>"
    FF = FreeFile
    Open MyManifestFile For Output As #FF
        Print #FF, XPStr
    Close #FF
    JustMade = True
Written:
    'SetAttr MyManifestFile, 34
    'the above line would have hidden the manifest, but I'd like my user to be able to delete that crap so i'll cancel that feature
    Dim iccex As tagInitCommonControlsEx
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_USEREX_CLASSES
    End With
    InitCommonControlsEx iccex
    XPVB = (Err.Number = 0)
    On Error GoTo 0
End Function

Public Function MyManifestFile() As String
    On Error Resume Next
    MyManifestFile = FindPath(App.Path, App.EXEName & ".exe.manifest")
End Function

Sub Main()
    On Error Resume Next
    If GetSet("MakeManifest", "0") = "1" Then XPVB
    frmMain.Show
    Dim K As String
    K = Replace(Command$, """", "") 'remove quotes
    If Len(K) > 0 Then 'if theres a command
        frmMain.DecideOnType K, TrimFileNameLOL(K) 'then try and load the file
    End If
End Sub
