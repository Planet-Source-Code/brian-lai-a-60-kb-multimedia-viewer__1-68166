VERSION 5.00
Begin VB.Form frmShowFileFilter 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Show File Type"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3975
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmShowFileFilter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3975
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.TextBox txtPattern 
      Height          =   315
      Left            =   120
      TabIndex        =   8
      Text            =   "*.*"
      Top             =   2160
      Width           =   3735
   End
   Begin VB.CommandButton btnExec 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2880
      TabIndex        =   5
      Top             =   2520
      Width           =   975
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "System"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Read Only"
      Height          =   255
      Index           =   3
      Left            =   2040
      TabIndex        =   3
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Normal"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   1815
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Hidden"
      Height          =   255
      Index           =   1
      Left            =   2040
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.CheckBox chkShow 
      Caption         =   "Archived"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin VB.Image imgTag 
      Height          =   360
      Left            =   120
      Top             =   120
      Width           =   360
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "Show only this kind of file..."
      Height          =   210
      Left            =   120
      TabIndex        =   7
      Top             =   1920
      Width           =   3390
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "Select the kind of files to show"
      Height          =   210
      Left            =   840
      TabIndex        =   6
      Top             =   240
      Width           =   2910
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmShowFileFilter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnExec_Click(Index As Integer)
    On Error Resume Next
    Unload Me
End Sub

Private Sub chkShow_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            frmMain.File.Archive = (chkShow(Index).Value = 1)
        Case 1
            frmMain.File.Hidden = (chkShow(Index).Value = 1)
        Case 2
            frmMain.File.Normal = (chkShow(Index).Value = 1)
        Case 3
            frmMain.File.ReadOnly = (chkShow(Index).Value = 1)
        Case 4
            frmMain.File.System = (chkShow(Index).Value = 1)
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
        
    SkinForm Me
    SkinFormEx Me

    chkShow(0).Value = IIf(frmMain.File.Archive, 1, 0)
    chkShow(1).Value = IIf(frmMain.File.Hidden, 1, 0)
    chkShow(2).Value = IIf(frmMain.File.Normal, 1, 0)
    chkShow(3).Value = IIf(frmMain.File.ReadOnly, 1, 0)
    chkShow(4).Value = IIf(frmMain.File.System, 1, 0)
    txtPattern.Text = FilterInUse
End Sub

Private Sub txtPattern_Change()
    On Error Resume Next
    With txtPattern
        If GetSet("Search_Fuzzy", "1") = "1" Then
                frmMain.File.Pattern = "*" & .Text & "*"
                frmMain.txtQuickFilter.Text = .Text
                frmMain.File.Pattern = Replace(frmMain.File.Pattern, "**", "*")
                frmMain.File.Pattern = Replace(frmMain.File.Pattern, ".*.", ".")
                SaveSet "File_Pattern", .Text
            Else
                If Len(.Text) = 0 Then .Text = "*.*"
                frmMain.File.Pattern = .Text
                frmMain.txtQuickFilter.Text = .Text
                frmMain.File.Pattern = Replace(frmMain.File.Pattern, "**", "*")
                frmMain.File.Pattern = Replace(frmMain.File.Pattern, ".*.", ".")
                SaveSet "File_Pattern", .Text
            End If
    End With
End Sub

