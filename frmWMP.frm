VERSION 5.00
Object = "{6BF52A50-394A-11D3-B153-00C04F79FAA6}#1.0#0"; "wmp.dll"
Begin VB.Form frmWMP 
   AutoRedraw      =   -1  'True
   Caption         =   "Media"
   ClientHeight    =   3255
   ClientLeft      =   60
   ClientTop       =   270
   ClientWidth     =   4695
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWMP.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   4695
   WindowState     =   2  '³Ì¤j¤Æ
   Begin WMPLibCtl.WindowsMediaPlayer WMP 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      URL             =   ""
      rate            =   1
      balance         =   0
      currentPosition =   0
      defaultFrame    =   ""
      playCount       =   999
      autoStart       =   -1  'True
      currentMarker   =   0
      invokeURLs      =   -1  'True
      baseURL         =   ""
      volume          =   50
      mute            =   0   'False
      uiMode          =   "full"
      stretchToFit    =   0   'False
      windowlessVideo =   0   'False
      enabled         =   -1  'True
      enableContextMenu=   -1  'True
      fullScreen      =   0   'False
      SAMIStyle       =   ""
      SAMILang        =   ""
      SAMIFilename    =   ""
      captioningID    =   ""
      enableErrorDialogs=   0   'False
      _cx             =   8070
      _cy             =   4683
   End
End
Attribute VB_Name = "frmWMP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    On Error Resume Next
    frmMain.titMedia.Visible = True
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    frmMain.titMedia.Visible = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
'    Mod32BitIcon.SetIcon Me.hwnd, "AAA"
    'settings
    Me.Icon = frmMain.Icon
    SkinForm Me
    SkinFormEx Me

    WMP.stretchToFit = GetSet("Media_Stretch", "1")
    WMP.uiMode = IIf(GetSet("Media_Controls", "1") = "1", "full", "none")
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    WMP.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Form_Deactivate
End Sub

Public Function LoadFile(TheFN As String) As Long
    On Error Resume Next
    WMP.URL = AddRecentItem(TheFN)
    Me.Caption = TrimFileNameLOL(TheFN)
    If GetSet("Sync_PSM", "0") = "1" Then 'sync MSN PSM
        Shell FindPath(App.Path, "PSMChanger.exe ") & App.ProductName & " - " & WMP.currentMedia.Name
    End If
    SaveSet "Media_Last", TheFN
    Me.Show
    SStatus Me.Name & " opened " & TheFN
End Function

Private Sub WMP_MediaError(ByVal pMediaObject As Object)
    On Error Resume Next
    SStatus "Error when trying to play " & WMP.URL
End Sub

