VERSION 5.00
Begin VB.Form frmPrefs 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Âù½u©T©w¹ï¸Ü¤è¶ô
   Caption         =   "Preferences"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrefs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '©ÒÄÝµøµ¡¤¤¥¡
   Begin VB.CommandButton btnUnloadMe 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5760
      TabIndex        =   10
      Top             =   5520
      Width           =   975
   End
   Begin VB.CommandButton btnOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4680
      TabIndex        =   2
      Top             =   5520
      Width           =   975
   End
   Begin VB.ListBox LstTab 
      Height          =   5340
      IntegralHeight  =   0   'False
      ItemData        =   "frmPrefs.frx":000C
      Left            =   120
      List            =   "frmPrefs.frx":0016
      TabIndex        =   1
      Top             =   120
      Width           =   1335
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   5295
      Index           =   1
      Left            =   1560
      ScaleHeight     =   5295
      ScaleWidth      =   5175
      TabIndex        =   6
      Top             =   120
      Width           =   5175
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   4
         Left            =   360
         TabIndex        =   25
         Text            =   "Text1"
         ToolTipText     =   "OpenOnStart,"
         Top             =   4920
         Width           =   4335
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Use Fuzzy Search"
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   23
         ToolTipText     =   "Search_Fuzzy,1"
         Top             =   1680
         Width           =   5175
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   3
         Left            =   360
         TabIndex        =   21
         Text            =   "Text1"
         ToolTipText     =   "Search_Provider_URL,http://www.google.com/search?hl=en&q=%s&btnG=Google+Search"
         Top             =   4200
         Width           =   4335
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   2
         Left            =   360
         TabIndex        =   19
         Text            =   "Text1"
         ToolTipText     =   "Search_Provider_Name,Google"
         Top             =   3480
         Width           =   4335
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   1
         Left            =   360
         TabIndex        =   17
         Text            =   "Text1"
         ToolTipText     =   "Filtre_String,Filter"
         Top             =   2760
         Width           =   4335
      End
      Begin VB.TextBox txtData 
         Height          =   315
         Index           =   0
         Left            =   360
         TabIndex        =   15
         Text            =   "Text1"
         ToolTipText     =   "Path Abbrev,...\"
         Top             =   1320
         Width           =   4335
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Sandbox mode"
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   13
         ToolTipText     =   "Sandbox,"
         Top             =   120
         Width           =   5175
      End
      Begin VB.CheckBox chkOpt 
         Caption         =   "Show full paths"
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   11
         ToolTipText     =   "ShowFullPaths,"
         Top             =   600
         Width           =   5175
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Show on Startup tag:"
         Height          =   210
         Index           =   7
         Left            =   360
         TabIndex        =   26
         Top             =   4680
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "e.g. Searches *Hello*.* when you enter Hello."
         ForeColor       =   &H00808080&
         Height          =   210
         Index           =   6
         Left            =   360
         TabIndex        =   24
         Top             =   1920
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Search Provider Search String:"
         Height          =   210
         Index           =   5
         Left            =   360
         TabIndex        =   22
         Top             =   3960
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Search Provider Name:"
         Height          =   210
         Index           =   4
         Left            =   360
         TabIndex        =   20
         Top             =   3240
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Text shown on the Filter box:"
         Height          =   210
         Index           =   3
         Left            =   360
         TabIndex        =   18
         Top             =   2520
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Abbreviation for shortened folder paths:"
         Height          =   210
         Index           =   2
         Left            =   360
         TabIndex        =   16
         Top             =   1080
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "Disallow writing of settings, including this one."
         ForeColor       =   &H00808080&
         Height          =   210
         Index           =   1
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   4620
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblNotification 
         AutoSize        =   -1  'True
         Caption         =   "on menus and forms"
         ForeColor       =   &H00808080&
         Height          =   210
         Index           =   0
         Left            =   360
         TabIndex        =   12
         Top             =   840
         Width           =   4620
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picTabSwitch 
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   5295
      Index           =   0
      Left            =   1560
      ScaleHeight     =   5295
      ScaleWidth      =   5175
      TabIndex        =   0
      Top             =   120
      Width           =   5175
      Begin VB.CheckBox chkOpt 
         Caption         =   "Make on startup"
         Height          =   255
         Index           =   3
         Left            =   2400
         TabIndex        =   28
         ToolTipText     =   "MakeManifest,0"
         Top             =   4380
         Width           =   5175
      End
      Begin VB.CommandButton btnShellINI 
         Caption         =   "&Edit INI..."
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   4800
         Width           =   2175
      End
      Begin VB.ListBox LstCredits 
         Height          =   1575
         IntegralHeight  =   0   'False
         ItemData        =   "frmPrefs.frx":002A
         Left            =   120
         List            =   "frmPrefs.frx":0043
         TabIndex        =   7
         Top             =   2640
         Width           =   4935
      End
      Begin VB.CommandButton btnWriteXPVS 
         Caption         =   "Make Manifest"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "List of beta testers: (thank you so much!)"
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   3465
      End
      Begin VB.Image imgLogo 
         Height          =   645
         Left            =   120
         Top             =   120
         Width           =   2190
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         X1              =   120
         X2              =   5040
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label lblProdVer 
         BackStyle       =   0  '³z©ú
         Caption         =   "Version "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   960
         Width           =   4935
      End
      Begin VB.Label lblProdDes 
         BackStyle       =   0  '³z©ú
         Caption         =   "Description"
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   4935
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '³z©ú
      Caption         =   "Apology: I got codes from too many people - not all have been credited in the source."
      Height          =   420
      Left            =   120
      TabIndex        =   27
      Top             =   5520
      Visible         =   0   'False
      Width           =   4455
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmPrefs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim InaSub As Boolean

Private Sub btnOK_Click()
    On Error Resume Next
    Dim i As Integer
    For i = 0 To chkOpt.UBound Step 1 'Save Settings
        SaveSet GetString(chkOpt(i).ToolTipText), Str$(chkOpt(i).Value)
    Next
    For i = 0 To txtData.UBound Step 1 'Save Settings
        SaveSet GetString(txtData(i).ToolTipText), txtData(i).Text
    Next
    
    Unload Me
End Sub

Private Sub btnShellINI_Click()
    On Error Resume Next
    Shell "notepad " & SettingsFile, vbNormalFocus
End Sub

Private Sub btnUnloadMe_Click()
    On Error Resume Next
    Unload Me
End Sub

Private Sub btnWriteXPVS_Click()
    On Error Resume Next
'    MsgBox "The Manifest module is not included in this program.", vbExclamation
    'If MsgBox("This function will write the manifest file again to show the Windows XP Visual Styles if applicable.", _
    vbYesNo + vbQuestion) = vbNo Then Exit Sub
    XPVB
    MsgBox "Manifest Written. Please restart " & App.ProductName & " to see effect.", vbInformation
End Sub

Private Sub chkOpt_Click(Index As Integer)
    On Error Resume Next
    If InaSub Then Exit Sub
    Select Case Index
        Case 3
            If chkOpt(Index).Value = 1 Then
                DSA 9
            End If
    End Select
End Sub

Private Sub Form_Load()
    On Error Resume Next
    InaSub = True
    lblProdVer.Caption = App.ProductName & " " & App.Major & "." & App.Minor & "." & App.Revision
    lblProdDes.Caption = App.ProductName & " " & MyVer & ", all rights reserved by Thinc." & vbCrLf & _
    "Made by Brian Lai" & vbCrLf & SoftwareHomePage
        
    SkinForm Me
    SkinFormEx Me

    For i = 0 To chkOpt.UBound Step 1 'Load Settings
        chkOpt(i).Value = GetSet(GetString(chkOpt(i).ToolTipText), GetString(chkOpt(i).ToolTipText, 1))
    Next
    For i = 0 To txtData.UBound Step 1 'Load Settings
        txtData(i).Text = GetSet(GetString(txtData(i).ToolTipText), GetString(txtData(i).ToolTipText, 1))
    Next
    InaSub = False
End Sub

Private Sub LstTab_Click()
    On Error Resume Next
    picTabSwitch(LstTab.ListIndex).ZOrder 0
End Sub

Public Function GoToTab(Index As Integer)
    On Error Resume Next
    picTabSwitch(Index).ZOrder 0
End Function
