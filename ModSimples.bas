Attribute VB_Name = "ModGeneric"
Option Explicit

Public Declare Function GetAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpSpec As String) As Long
Public Declare Function SetAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpSpec As String, ByVal dwAttributes As Long) As Long

Public Enum vbFileAttributes
  vbNormal = 0         ' Normal
  vbReadOnly = 1       ' Read-only
  vbHidden = 2         ' Hidden
  vbSystem = 4         ' System file
  vbVolume = 8         ' Volume label
  vbDirectory = 16     ' Directory or folder
  vbArchive = 32       ' File has changed since last backup
  vbTemporary = &H100  ' 256
  vbCompressed = &H800 ' 2048
End Enum

Public Function GetAttrib(sFileSpec As String, ByVal Attrib As vbFileAttributes) As Boolean
  If (LenB(sFileSpec) <> 0) Then
    GetAttrib = (GetAttributes(sFileSpec) And Attrib) = Attrib
  End If
End Function

Public Sub SetAttrib(sFileSpec As String, ByVal Attrib As vbFileAttributes, Optional fTurnOff As Boolean)
  If (LenB(sFileSpec) <> 0) Then
    If (Attrib = vbNormal) Then
      SetAttributes sFileSpec, vbNormal
    ElseIf fTurnOff Then
      SetAttributes sFileSpec, GetAttributes(sFileSpec) And (Not Attrib)
    Else
      SetAttributes sFileSpec, GetAttributes(sFileSpec) Or Attrib
    End If
  End If
End Sub

Function ControlNum(YourVal As Long, Optional MinVal As Long = 0, Optional MaxVal As Long = 100) As Long
    On Error Resume Next
    If YourVal < MinVal Then YourVal = MinVal
    If YourVal > MaxVal Then YourVal = MaxVal
    ControlNum = YourVal
End Function

Public Function FindPath(Parent As String, Optional Child As String, Optional Divider As String = "\") As String
    On Error Resume Next
    If Right$(Parent, 1) = Divider Then Parent = Left$(Parent, Len(Parent) - 1)
    If Left$(Child, 1) = Divider Then Child = Mid$(Child, 2)
    FindPath = Parent & Divider & Child
End Function

Public Function GetString(Which As String, Optional SectionNo As Long = 0, Optional Delimiter As String = ",") As String
    On Error Resume Next
    Dim Arr() As String
    Arr = Split(Which, Delimiter)
    GetString = Arr(SectionNo)
End Function

Public Sub SStatus(Optional What As String = "Ready")
    On Error Resume Next
    With frmMain
        If .lblStatus.Caption = What Then Exit Sub
        .lblStatus.Caption = What
    End With
End Sub

Public Sub SProgress(Value As Long, Optional ValueMin As Long = 0, Optional ValueMax As Long = 100)
    On Error GoTo errrr
    Dim B As Double
    With frmMain
        If Value <= ValueMin Or Value > ValueMax Then
            .picProgress.Visible = False
        Else
            .picProgress.Visible = True
            .picProgress.Height = .ImgProgress.Height
            .ImgProgressVal.Height = .ImgProgress.Height
            .ImgProgress.Width = .picProgress.Width 'fill with image
            B = IIf(ValueMax <= 0, 0, Value / ValueMax)
            .ImgProgressVal.Width = IIf(B = 0, 0, (.ImgProgress.Width - .ImgProgressVal.Left) * B)
        End If
        'this part is for SStatus
        If .picProgress.Visible Then
            .lblStatus.Left = .picProgress.Left + .picProgress.Width + 60
        Else
            .lblStatus.Left = 30
        End If
        '/this part is for SStatus
    End With
    Exit Sub
errrr:
    Debug.Print Err.Description
End Sub

Public Function TrimFileNameLOL(FromWhat As String, Optional ForceLong As Boolean = False, _
                                                                            Optional AddDotDotDot As Boolean = False) As String
    On Error Resume Next
    If GetSet("ShowFullPaths") = "1" Or ForceLong Then
        TrimFileNameLOL = FromWhat
    Else
        TrimFileNameLOL = Right$(FromWhat, Len(FromWhat) - InStrRev(FromWhat, "\"))
        If AddDotDotDot Then
            If Right$(TrimFileNameLOL, 2) <> ":\" Then 'only if it's not a drive
                TrimFileNameLOL = GetSet("Path_Abbrev", PathAbbrev) & TrimFileNameLOL
            End If
        End If
    End If
End Function

Public Sub TXTFileSave(Text As String, filepath As String)
    On Error Resume Next
    Dim F As Integer
    F = FreeFile
    Open filepath For Binary As #F
    Put #F, , Text
    Close #F
    Exit Sub
End Sub

Public Function MyVer() As String
    On Error Resume Next
    Dim Buffer2 As String
    Dim PreVer As Integer
    PreVer = App.Minor
    If App.Revision >= 1 Then
        PreVer = PreVer + 1
        Buffer2 = Trim$(Str$(PreVer) & " BETA")
    Else
        Buffer2 = Trim$(Str$(PreVer))
    End If
    MyVer = "V." & App.Major & "." & Buffer2
End Function


Public Function FileExists(TheFN As String) As Boolean
    'does not work for hidden or whtever objects.
    'edited from copied code. his remarks:
    '#          I wish that it will help sumebody              #
    '#     I think this is one of the easiest way to do it     #
    Dim Var1 As String       'Variable for this module.
    On Error GoTo NotThere       'Simulate the occurrence of an error.
    Var1 = Dir$(TheFN) 'send back a string value.
    FileExists = (Var1 <> "")        'True = 1
NotThere:                            'The error reference.
    If Err = 53 Then Resume Next 'If the Simulate Error occure then will resume next.
End Function


