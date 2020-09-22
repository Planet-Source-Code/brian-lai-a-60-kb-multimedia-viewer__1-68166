VERSION 5.00
Begin VB.Form frmBrwFav 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Favorites"
   ClientHeight    =   3855
   ClientLeft      =   4950
   ClientTop       =   4035
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBrwFav.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton btnFavExec 
      Caption         =   "Remove"
      Height          =   375
      Index           =   3
      Left            =   960
      TabIndex        =   4
      ToolTipText     =   "Delete selected bookmark"
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton btnFavExec 
      Caption         =   "Add"
      Height          =   375
      Index           =   2
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "Add this page"
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton btnFavExec 
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   4560
      TabIndex        =   2
      Top             =   3480
      Width           =   975
   End
   Begin VB.CommandButton btnFavExec 
      Caption         =   "Open"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   3600
      TabIndex        =   1
      Top             =   3480
      Width           =   975
   End
   Begin VB.FileListBox filFav 
      Height          =   2820
      Left            =   0
      Pattern         =   "*.url"
      TabIndex        =   0
      Top             =   0
      Width           =   5535
   End
   Begin VB.Label lblFN 
      BackStyle       =   0  '³z©ú
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   570
      Left            =   0
      TabIndex        =   5
      Top             =   2880
      UseMnemonic     =   0   'False
      Width           =   5535
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmBrwFav"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    On Error Resume Next
    Me.Move frmMain.Left + (frmMain.Width - Me.Width) / 2, frmMain.Top + (frmMain.Height - Me.Height) / 2 'have to do this... the center of the drawn form was on a button. wtf
    filFav.Path = FavsPath
    filFav.Refresh
        
    SkinForm Me
    SkinFormEx Me

    filFav.ListIndex = 0
    filFav_Click 'activate the first one
End Sub

Private Sub btnFavExec_Click(Index As Integer)
    On Error Resume Next
    If Index = 0 Then
        FavStr = FavAddy(FindPath(filFav.Path, filFav.filename))
    ElseIf Index = 2 Then
        Dim A As String
        A = InputBox("Enter Name of your shortcut:", , Replace(AF.BRW.LocationName, "/", "_"))
        If Len(A) > 0 Then
            WriteINI "InternetShortcut", "URL", AF.BRW.LocationURL, FindPath(FavsPath, A & ".url")
        End If
    ElseIf Index = 3 Then
        If MsgBox("Are you sure you want to delete the bookmark " & vbCrLf & filFav.filename & "?", vbYesNo + vbQuestion) = vbYes Then
            Kill FindPath(filFav.Path, filFav.filename)
            filFav.Refresh
        End If
    End If
    If Index < 2 Then Unload Me
End Sub

Public Function FavAddy(WhichFile As String) As String
    On Error Resume Next
    FavAddy = AF.FavAddy(WhichFile) 'wrap
End Function

Private Sub filFav_Click()
    On Error Resume Next
    lblFN.Caption = filFav.filename & vbCrLf & FavAddy(FindPath(filFav.Path, filFav.filename))
    'lblURL.Caption = FavAddy(FindPath(filFav.Path, filFav.filename))
End Sub

Private Sub filFav_DblClick()
    On Error Resume Next
    btnFavExec_Click 0
End Sub

Private Sub lblFN_DblClick()
    On Error Resume Next
    Clipboard.SetText FavAddy(FindPath(filFav.Path, filFav.filename))
    Beep
End Sub
