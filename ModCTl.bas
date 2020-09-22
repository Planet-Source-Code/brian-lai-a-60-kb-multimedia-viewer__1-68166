Attribute VB_Name = "ModCTl"
Option Explicit

Public FavStr As String

Public Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nSize As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Function GetWindowLong& Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long)
Public Declare Function lstrcat Lib "kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetWindowLong& Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long)
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Declare Function SHBrowseForFolder Lib "SHELL32" (lpBI As BrowseInfo) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function SHGetPathFromIDList Lib "SHELL32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long
Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

Public Const DefaultSearchURL As String = "http://www.google.com/search?hl=en&q=%s&btnG=Google+Search"
Public Const SoftwareHomePage As String = "http://thinc.myvnc.com"
Public Const DefaultSearchAgent As String = "Google"
Public Const DefaultFilterString As String = "Filter"
Public Const DefaultTmpFileName As String = "ProFile.tmp"
Public Const PathAbbrev As String = "..\"
Public Const LB_ITEMFROMPOINT = &H1A9
Public Const BIF_RETURNONLYFSDIRS = 1
Public Const BIF_DONTGOBELOWDOMAIN = 2
Public Const BIF_BROWSEFORCOMPUTER = &H1000
Public Const MAX_PATH = 260

Private Type BrowseInfo
    hwndOwner As Long
    pIDLRoot As Long
    pszDisplayName As Long
    lpszTitle As Long
    ulFlags As Long
    lpfnCallback As Long
    lParam As Long
    iImage As Long
End Type

Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAnyOperationsAborted As Long
    hNameMappings As Long
    lpszProgressTitle As Long
End Type


Public Function UserName() As String
On Error Resume Next
    Dim lpBuffer As String
    Dim J
    lpBuffer = Space$(255)
    GetUserName lpBuffer, Len(lpBuffer)
        J = InStr(lpBuffer, Chr$(0))
    If J > 0 Then UserName = Left$(lpBuffer, J - 1)
End Function

Public Sub SkinFormEx(Which As Form)
    On Error Resume Next
    Dim A As Control
    For Each A In Which
        BTFlat A
        CtlFlat A
    Next 'there arent any Flat Buttons
    FormFlat Which
End Sub

Public Sub BTFlat(BT As CommandButton)
    On Error Resume Next
    If FileExists(MyManifestFile) = True And JustMade = False Then Exit Sub 'flatten if no manifest
    If GetWindowLong&(BT.hwnd, -16) And &H8000& Then Exit Sub
    SetWindowLong BT.hwnd, -16, GetWindowLong&(BT.hwnd, -16) Or &H8000&
    BT.Refresh
End Sub

Public Sub CtlFlat(CL As Control)
    On Error Resume Next
    If FileExists(MyManifestFile) = True And JustMade = False Then Exit Sub  'generate if no manifest
    CL.Appearance = 0 'flat
    CL.BackColor = frmMain.picBrw.BackColor  'for cham buttons, and they change backcolor to the same as the container
    'for cham buttons only
        CL.ColorScheme = 2
        CL.BackOver = frmMain.picBrw.BackColor
    'for cham
    If Err.Number <> 0 Then Debug.Print Err.Description
End Sub

Public Sub FormFlat(Which As Form)
    On Error Resume Next
    If FileExists(MyManifestFile) = True And JustMade = False Then Exit Sub  'flatten if no manifest
    Which.Appearance = 0
End Sub

Public Sub LoadSearchProvider()
    On Error Resume Next
    frmMain.txtSearch.Text = GetSet("Search_Provider_Name", DefaultSearchAgent) & "..."
    frmMain.txtSearch.Tag = GetSet("Search_Provider_URL", DefaultSearchURL)
End Sub

Function FileText(ByVal FN As String) As String
    Dim handle As Integer
    If Len(Dir$(FN)) = 0 Then FileText = ""
    handle = FreeFile
    Open FN For Binary As #handle
    ' read the string and close the file
    FileText = Space$(LOF(handle))
    Get #handle, , FileText
    Close #handle
End Function

Private Function GetTempDir() As String
   Dim nSize As Long
   Dim tmp As String
   tmp = Space$(MAX_PATH)
   nSize = Len(tmp)
   Call GetTempPath(nSize, tmp)
   GetTempDir = TrimNull(tmp)
End Function

Private Function TrimNull(item As String)
   Dim pos As Long
   pos = InStr(item, vbNullChar)
    TrimNull = IIf(pos, Left$(item, pos - 1), item)
End Function

Public Function BrowseForFolder(Owner As Long, Optional szTitle As String = "Select Folder...") As String
    Dim lpIDList As Long
    Dim sBuffer As String
    Dim tBrowseInfo As BrowseInfo

    With tBrowseInfo
        .hwndOwner = Owner
        .lpszTitle = lstrcat(szTitle, "")
        .ulFlags = BIF_RETURNONLYFSDIRS + BIF_DONTGOBELOWDOMAIN
    End With
    lpIDList = SHBrowseForFolder(tBrowseInfo)

    If (lpIDList) Then
        sBuffer = Space(MAX_PATH)
        SHGetPathFromIDList lpIDList, sBuffer
        sBuffer = Left(sBuffer, InStr(sBuffer, vbNullChar) - 1)
        BrowseForFolder = sBuffer
    End If
End Function

Public Function AF() As Form
    On Error Resume Next
    Set AF = frmMain.ActiveForm 'returns active form... in a shorter syntax in the main program.
End Function

Public Function FilterInUse() As String
    On Error Resume Next
    FilterInUse = frmMain.txtQuickFilter.Text
End Function

Public Function DisplayFilter() As String
    On Error Resume Next
    If GetSet("Search_Fuzzy", "1") = "1" Then
        DisplayFilter = GetSet("File_Pattern")
        Do Until InStr(1, DisplayFilter, "**") < 1
            DisplayFilter = Replace(DisplayFilter, "**", "*")
            DisplayFilter = Replace(DisplayFilter, ".*.", ".")
        Loop
        DisplayFilter = Mid$(DisplayFilter, 2, Len(DisplayFilter) - 2) 'trim the autocodes
        If DisplayFilter = "." Then DisplayFilter = GetSet("Filtre_String", DefaultFilterString) & "..."
    Else
        DisplayFilter = GetSet("File_Pattern", "*.*")
    End If
End Function

Public Function DownloadFile(URL As String, Optional SaveAsFile As String) As String
    On Error Resume Next
    If Len(SaveAsFile) = 0 Then SaveAsFile = FindPath(GetTempDir, DefaultTmpFileName)
    URLDownloadToFile 0, URL, SaveAsFile, 0, 0
    DownloadFile = SaveAsFile
End Function

Public Sub LoadRecents()
    On Error Resume Next
    Dim i As Integer
    Dim J As String
    With frmMain
        For i = 0 To 9 Step 1
            J = GetSet("Recent" & i)
            .titLERecentFilesArray(i).Caption = TrimFileNameLOL(J)
            .titLERecentFilesArray(i).Tag = J
            .titLERecentFilesArray(i).Visible = (.titLERecentFilesArray(i).Tag <> "")
        Next
    End With
End Sub

Public Function AddRecentItem(WhatFileName As String) As String
    On Error Resume Next
    Dim i As Integer
    Dim A As String
        For i = 9 To 0 Step -1
            If GetSet("Recent" & i) = WhatFileName Then
                AddRecentItem = WhatFileName
                Exit Function
            End If
        Next
        For i = 9 To 1 Step -1 'end with 1! the first one is discarded ah mah...
            A = GetSet("Recent" & i - 1)
            If Len(A) > 0 Then
                SaveSet "Recent" & i, A
            End If
        Next i
        SaveSet "Recent0", WhatFileName
        LoadRecents
        AddRecentItem = WhatFileName
End Function

Public Sub LoadRecentFolders()
    On Error Resume Next
    Dim i As Integer
    Dim J As String
    With frmMain
        For i = 0 To 9 Step 1
            J = GetSet("RecentF" & i)
            .titLERecentFoldersArray(i).Caption = TrimFileNameLOL(J, , True)
            .titLERecentFoldersArray(i).Tag = J
            .titLERecentFoldersArray(i).Visible = (.titLERecentFoldersArray(i).Tag <> "")
        Next
    End With
End Sub

Public Function AddRecentFolder(WhatFolderName As String) As String
    On Error Resume Next
    Dim A As String
    Dim i As Integer ', J As Integer
        For i = 9 To 0 Step -1
            If GetSet("RecentF" & i) = WhatFolderName Then
                AddRecentFolder = WhatFolderName
                Exit Function
            End If
        Next
        For i = 9 To 1 Step -1 'end with 1! the first one is discarded ah mah...
            A = GetSet("RecentF" & i - 1)
            If Len(A) > 0 Then
                SaveSet "RecentF" & i, A
            End If
        Next i
        SaveSet "RecentF0", WhatFolderName
        LoadRecentFolders
        AddRecentFolder = WhatFolderName
End Function

Public Function MyMsgBox(ShowonForm As String, SaveNum As Integer, Optional ShowOnWindow As String, Optional HideCHK As Boolean) As Long
    On Error Resume Next
    MyMsgBox = frmInputMsg.MyMsgBoxEx(ShowonForm, SaveNum, ShowOnWindow, HideCHK) 'wrapper
End Function

Public Function ParsedAddy(WhatNow As String) As String
    On Error Resume Next
    Dim CMD As String, SRC As String
    If InStr(1, WhatNow, " ") <= 0 Then
        ParsedAddy = WhatNow
    Else
        CMD = LCase$(Trim$(Left$(WhatNow, InStr(1, WhatNow, " "))))
        SRC = Mid$(WhatNow, InStr(1, WhatNow, " ") + 1)
        Select Case CMD
            Case "g", "google"
                ParsedAddy = "http://www.google.com/search?hl=en&q=%s&btnG=Google+Search"
            Case "vb", "vb6"
                ParsedAddy = "http://www.planet-source-code.com/vb/scripts/BrowseCategoryOrSearchResults.asp?optSort=Alphabetical&txtCriteria=%s&blnWorldDropDownUsed=TRUE&txtMaxNumberOfEntriesPerPage=50&blnResetAllVariables=TRUE&lngWId=1"
            Case "i", "image"
                ParsedAddy = "http://images.google.com/images?q=%s"
            Case "u", "youtube"
                ParsedAddy = "http://www.youtube.com/results?search_query=%s"
        End Select
        ParsedAddy = Replace(ParsedAddy, "%s", SRC)
    End If
End Function

Public Function FavsPath() As String
    On Error Resume Next
    FavsPath = "C:\Documents and Settings\" & UserName & "\Favorites"
End Function

Public Function OnTop(TheHwnd As Long, TrueOrFalse As Boolean)
    On Error Resume Next
    SetWindowPos TheHwnd, IIf(TrueOrFalse, -1, -2), 0, 0, 0, 0, 3 '&H1 Or &H10 Or &H2 Or &H40
End Function

Public Function DSA(Index As Integer) 'dont show again collection
    Dim K As String
    Select Case Index
        Case 1
            K = "You are in sandbox mode - this setting will not be changed. You will need to Edit the INI manually in the preferences window."
        Case 2
            K = "To start using " & App.ProductName & ", please navigate to a file with the browser bar."
        Case 3
            K = App.ProductName & " will run this code only if the component ""ESE.exe"" is available." & vbCrLf & vbCrLf & "You can get a copy of the component at " & SoftwareHomePage & "."
        Case 4
            K = App.ProductName & " will download this file and may stop responding while downloading a large file."
        Case 5
            K = "The sandbox mode stops your settings file from being changed or written into."
        Case 6
            K = "This will also change your start up form tag to Media."
        Case 7
            K = App.ProductName & " will run this code only if the component ""PSMChanger.exe"" is available." & vbCrLf & vbCrLf & "You can get a copy of the component at " & SoftwareHomePage & "."
        Case 8 'file info diag
            DoEvents
        Case 9
            K = App.ProductName & " may not able to start after this." & vbCrLf & "To reset this:" & vbCrLf & "- delete " & App.EXEName & ".exe.manifest from the directory." & vbCrLf & "- come back to this window to uncheck this checkbox."
        Case 10
            K = App.ProductName & " will run this code only if the component ""TEExt.exe"" is available." & vbCrLf & vbCrLf & "You can get a copy of the component at " & SoftwareHomePage & "."
    End Select
    MyMsgBox K, Index
End Function
