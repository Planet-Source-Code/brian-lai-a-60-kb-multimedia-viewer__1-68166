Attribute VB_Name = "ModSkin"
Option Explicit
Private Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal sSectionName As String, ByVal sReturnedString As String, ByVal lSize As Long, ByVal sFileName As String) As Long
'No need for other INI readers

Function SkinForm(WhichForm As Form, Optional FromINIFile As String)
    On Error Resume Next
    Dim MyKeys As String * 1000 'I set this 255 from the original 1000
    Dim EachElement As Variant, EachKey As Variant
    Dim CtlName As String, CtlProp As String, CtlPropVal As String
    Dim ItemIdx As Integer
    If Len(FromINIFile) = 0 Then FromINIFile = GetSet("SkinFile", FindPath(App.Path, "skin.ini"))
    FromINIFile = Replace(FromINIFile, "{app}", App.Path) 'Converts to local paths
    If Dir(FromINIFile) = "" Then Exit Function 'If there's no such file, who cares?
    'SkinForm uses GetPrivateProfileSection
    GetPrivateProfileSection WhichForm.Name, MyKeys, 1000, FromINIFile
    EachKey = Split(MyKeys, Chr(0))
    For Each EachElement In EachKey
        If EachElement = "" Then
            Exit For
        End If
        'EachElement is in Label1 BackColor=255 form right now, so split with GetPrivString
        CtlName = GetPrivString(GetPrivString(EachElement, 0, "="), 0, " ") 'This gets the control name
        ItemIdx = Val(Mid$(CtlName, InStr(1, CtlName, "(") + 1, (InStrRev(CtlName, ")") - (InStr(1, CtlName, "(") + 1))))
        'ItemIdx is calculated by a crazy length of code
        CtlProp = GetPrivString(GetPrivString(EachElement, 0, "="), 1, " ") 'This gets the property, e.g. BackColor
        CtlPropVal = GetPrivString(EachElement, 1, "=") 'This gets the value of that property
        CtlPropVal = Replace(CtlPropVal, "{a}", App.Path) 'Converts to local paths
        CtlPropVal = Replace(CtlPropVal, "{app}", App.Path) 'Converts to local paths
        CtlPropVal = Replace(CtlPropVal, "{s}", Left$(FromINIFile, InStrRev(FromINIFile, "\") - 1)) 'Converts to local paths
        CtlPropVal = Replace(CtlPropVal, "{skin}", Left$(FromINIFile, InStrRev(FromINIFile, "\") - 1)) 'Converts to local paths
        Debug.Print CtlName, ItemIdx, CtlProp, CtlPropVal
        Select Case UCase$(CtlProp)
            Case "BACKCOLOR", "BC"
                If UCase$(CtlProp) = "BC" Then CtlProp = "backcolor" 'restore shortened var
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).item(ItemIdx).BackColor = Val(CtlPropVal)
                Else
                    WhichForm.Controls(CtlName).BackColor = Val(CtlPropVal)
                End If
            Case "FORECOLOR", "FC"
                If UCase$(CtlProp) = "FC" Then CtlProp = "forecolor" 'restore shortened var
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).item(ItemIdx).ForeColor = Val(CtlPropVal)
                Else
                    WhichForm.Controls(CtlName).ForeColor = Val(CtlPropVal)
                End If
            Case "PICTURE", "PIC"
                If UCase$(CtlProp) = "PIC" Then CtlProp = "picture" 'restore shortened var
                If InStr(1, CtlPropVal, "/") > 0 Then 'if this is not a local thing
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    CtlPropVal = DownloadFile(CtlPropVal) 'downloads the file from the net and returns path
                End If
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).item(ItemIdx).Picture = LoadPicture(CtlPropVal)
                Else
                    WhichForm.Controls(CtlName).Picture = LoadPicture(CtlPropVal)
                End If
            Case "PICTURENORMAL", "PN"
                If UCase$(CtlProp) = "PN" Then CtlProp = "picturenormal" 'restore shortened var
                If InStr(1, CtlPropVal, "/") > 0 Then 'if this is not a local thing
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    CtlPropVal = DownloadFile(CtlPropVal) 'downloads the file from the net and returns path
                End If
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).item(ItemIdx).PictureNormal = LoadPicture(CtlPropVal)
                Else
                    WhichForm.Controls(CtlName).PictureNormal = LoadPicture(CtlPropVal)
                End If
            Case "CAPTION", "CPN"
                If UCase$(CtlProp) = "CPN" Then CtlProp = "caption" 'restore shortened var
                If InStr(1, CtlName, "(") > 0 Then 'If this is an array...
                    CtlName = Mid$(CtlName, 1, InStr(1, CtlName, "(") - 1) 'Removes the index from name
                    WhichForm.Controls(CtlName).item(ItemIdx).Caption = CtlPropVal
                Else
                    WhichForm.Controls(CtlName).Caption = CtlPropVal
                End If
            Case "ICON"
        End Select
    Next
End Function

Private Function GetPrivString(Which As Variant, Optional SectionNo As Long = 0, Optional Delimiter As String = ",") As String
    'This GetPrivString is only for use in SkinForm because the source is a Variant there...
    On Error Resume Next
    Dim Arr() As String
    Arr = Split(Which, Delimiter)
    GetPrivString = Arr(SectionNo)
End Function


