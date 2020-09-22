VERSION 5.00
Begin VB.Form frmTXT 
   AutoRedraw      =   -1  'True
   Caption         =   "Text"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   420
   ClientWidth     =   4635
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTXT.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4635
   WindowState     =   2  '³Ì¤j¤Æ
   Begin VB.TextBox txtBox 
      Height          =   2775
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  '««ª½±²¶b
      TabIndex        =   0
      Top             =   0
      Width           =   4335
   End
End
Attribute VB_Name = "frmTXT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Activate()
    On Error Resume Next
    frmMain.titText.Visible = True
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    frmMain.titText.Visible = False
End Sub

Private Sub Form_Load()
    On Error Resume Next
'    Mod32BitIcon.SetIcon Me.hwnd, "AAA"
    Me.Icon = frmMain.Icon
    SkinForm Me
    SkinFormEx Me

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If Right$(Me.Caption, 1) = "*" Then 'unsaved eh
        Dim A As VbMsgBoxResult
        A = MsgBox("Do you want to save this file first?", vbYesNoCancel + vbQuestion)
        Select Case A
            Case vbYes
                frmMain.titTextFileSave_Click
            Case vbNo
                'do nothing?
            Case vbCancel
                Cancel = 1
        End Select
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    txtBox.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub

Public Function LoadFile(TheFN As String) As Long
    On Error Resume Next
    If FileLen(TheFN) > 64000 Then 'use for big files
        txtBox.Text = FileText(AddRecentItem(TheFN))
    Else
        Dim F As Integer
        Dim tmp As String, K As String
        F = FreeFile
        Open TheFN For Input As #F
            Do
                Line Input #F, tmp
                K = K & tmp & vbCrLf
            Loop Until EOF(F)
        Close #F
        txtBox.Text = K
    End If
    Me.Tag = TheFN
    Me.Caption = TrimFileNameLOL(TheFN)
    Me.Show
    SStatus Me.Name & " opened " & TheFN
End Function

Public Sub ChangeFont()
    On Error Resume Next
    Dim Response As VbMsgBoxResult
    With txtBox
        ShowFont
        .FontName = SelectFont.mFontName
        .FontSize = SelectFont.mFontsize
        .FontBold = SelectFont.mBold
        .FontItalic = SelectFont.mItalic
        .FontStrikethru = SelectFont.mStrikethru
        .FontUnderline = SelectFont.mUnderline
        .ForeColor = SelectFont.mFontColor
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Form_Deactivate
End Sub

Private Sub txtBox_Change()
    On Error Resume Next
    If Right$(Me.Caption, 1) <> "*" Then Me.Caption = Me.Caption & "*" 'state of change
End Sub
