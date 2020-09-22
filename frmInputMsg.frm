VERSION 5.00
Begin VB.Form frmInputMsg 
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   " "
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2295
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.CommandButton BTN 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   1800
      Width           =   1095
   End
   Begin VB.CheckBox CHK 
      Caption         =   "Don't show this again"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   1920
      Width           =   3015
   End
   Begin VB.Image IMG 
      Height          =   480
      Left            =   120
      Top             =   120
      Width           =   480
   End
   Begin VB.Label LBL 
      BackStyle       =   0  '³z©ú
      Caption         =   "Label1"
      Height          =   1530
      Left            =   720
      TabIndex        =   1
      Top             =   120
      Width           =   3615
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmInputMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Function MyMsgBoxEx(ShowonForm As String, SaveNum As Integer, Optional MyCaption As String, Optional HideCheckBox As Boolean) As Long
    On Error Resume Next
    If GetSet("DSA" & SaveNum) = "1" Then 'if this message is set not to show again
        MyMsgBoxEx = 1
        Exit Function 'then exit
    End If
    LBL.Caption = ShowonForm
    If HideCheckBox Then CHK.Visible = False
    Me.Caption = IIf(Len(MyCaption) > 0, MyCaption, App.ProductName)
    Me.Tag = SaveNum
    Me.Show 1
    MyMsgBoxEx = 1
End Function

Private Sub BTN_Click()
    On Error Resume Next
    SaveSet "DSA" & Me.Tag, CHK.Value 'save if show again
    Unload Me
End Sub

Private Sub Form_Load()
    On Error Resume Next
        
    SkinForm Me
    SkinFormEx Me

End Sub

