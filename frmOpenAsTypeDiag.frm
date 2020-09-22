VERSION 5.00
Begin VB.Form OpenFileDlg 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Open File"
   ClientHeight    =   3855
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3855
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmOpenAsTypeDiag.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3855
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.PictureBox Picture2 
      Height          =   3495
      Left            =   3840
      ScaleHeight     =   3435
      ScaleWidth      =   435
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   495
      Begin VB.Image ImgArray 
         Height          =   480
         Index           =   4
         Left            =   0
         Top             =   1920
         Width           =   480
      End
      Begin VB.Image ImgArray 
         Height          =   480
         Index           =   6
         Left            =   0
         Top             =   2880
         Width           =   480
      End
      Begin VB.Image ImgArray 
         Height          =   480
         Index           =   5
         Left            =   0
         Top             =   2400
         Width           =   480
      End
      Begin VB.Image ImgArray 
         Height          =   480
         Index           =   0
         Left            =   0
         Top             =   0
         Width           =   480
      End
      Begin VB.Image ImgArray 
         Height          =   480
         Index           =   1
         Left            =   0
         Top             =   480
         Width           =   480
      End
      Begin VB.Image ImgArray 
         Height          =   480
         Index           =   2
         Left            =   0
         Top             =   960
         Width           =   480
      End
      Begin VB.Image ImgArray 
         Height          =   480
         Index           =   3
         Left            =   0
         Top             =   1440
         Width           =   480
      End
   End
   Begin VB.ListBox lstFT 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2100
      IntegralHeight  =   0   'False
      ItemData        =   "frmOpenAsTypeDiag.frx":000C
      Left            =   120
      List            =   "frmOpenAsTypeDiag.frx":0025
      TabIndex        =   6
      Top             =   840
      Width           =   3615
   End
   Begin VB.CheckBox chkDontAsk 
      Caption         =   "Don't ask me again"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   3000
      Width           =   3615
   End
   Begin VB.CommandButton btnExec 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   3360
      Width           =   975
   End
   Begin VB.CommandButton btnExec 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2760
      TabIndex        =   2
      Top             =   3360
      Width           =   975
   End
   Begin VB.Image IMG 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblFormat 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "FormatCode"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   120
      Width           =   1965
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "What kind of file is this?"
      Height          =   210
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   2250
   End
End
Attribute VB_Name = "OpenFileDlg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MyChoice As Long
Dim MyHover As Integer

Public Function AsType(FileExtension As String, Optional FileNme As String, Optional IgnoreRMBFileExtFlag As Boolean) As Long
    On Error Resume Next
    Dim A As String
    A = Trim$(GetSet("PFT_" & UCase$(FileExtension), "-1"))
    If A <> "" And A <> "-1" Then 'say, -1 is the "undefined" number
        AsType = Val(Left$(A, 2))
        chkDontAsk.Value = 1
        lstFT.ListIndex = AsType
        If IgnoreRMBFileExtFlag Then GoTo ThereInstead 'ya ok...
        Exit Function
    Else
        chkDontAsk.Value = 0
    End If
ThereInstead:
    lblFormat.Caption = UCase$(FileExtension)
    Me.Tag = FileNme
    Me.Show 1
    AsType = MyChoice
End Function

Private Sub btnExec_Click(Index As Integer)
    On Error Resume Next
    Dim A As String
    If Index = 0 Then
        Dim i As Integer
        MyChoice = lstFT.ListIndex
        If chkDontAsk.Value = 1 Then
            A = Str(MyChoice)
            If A <> "-1" Then 'only make a new record if it's not "undefined"
                SaveSet "PFT_" & UCase$(lblFormat.Caption), A
            End If
        Else
            A = "-1"
        End If
    Else
        MyChoice = 99 'Ridiculous
    End If
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
        
    SkinForm Me
    SkinFormEx Me

End Sub

Private Sub lstFT_Click()
    On Error Resume Next
    MyChoice = lstFT.ListIndex
    If MyHover <> MyChoice Then
        IMG.Picture = ImgArray(MyChoice).Picture
        MyHover = MyChoice
    End If
End Sub

Private Sub lstFT_DblClick()
    On Error Resume Next
    btnExec_Click 0
End Sub

Private Sub lstFT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    lstFT_MouseMove 1, 0, 0, 0
End Sub

Private Sub lstFT_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 1 Then
        lstFT_Click
    End If
End Sub
