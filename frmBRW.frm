VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmBRW 
   AutoRedraw      =   -1  'True
   Caption         =   "Browser"
   ClientHeight    =   5025
   ClientLeft      =   3060
   ClientTop       =   3420
   ClientWidth     =   8745
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmBRW.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   8745
   ShowInTaskbar   =   0   'False
   WindowState     =   2  '³Ì¤j¤Æ
   Begin SHDocVwCtl.WebBrowser BRW 
      Height          =   2295
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   4680
      ExtentX         =   8255
      ExtentY         =   4048
      ViewMode        =   1
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   -1  'True
      NoClientEdge    =   -1  'True
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.PictureBox picAddress 
      Align           =   1  '¹ï»ôªí³æ¤W¤è
      BorderStyle     =   0  '¨S¦³®Ø½u
      BeginProperty Font 
         Name            =   "·s²Ó©úÅé"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   8745
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   8745
      Begin ProFile.chameleonButton btnBrw 
         Height          =   435
         Index           =   0
         Left            =   0
         TabIndex        =   4
         ToolTipText     =   "Back"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":000C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox cboAddress 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3960
         TabIndex        =   2
         Top             =   60
         Width           =   3795
      End
      Begin ProFile.chameleonButton btnBrw 
         Height          =   435
         Index           =   1
         Left            =   360
         TabIndex        =   5
         ToolTipText     =   "Forward"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":0028
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.chameleonButton btnBrw 
         Height          =   435
         Index           =   2
         Left            =   840
         TabIndex        =   6
         ToolTipText     =   "Refresh"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":0044
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.chameleonButton btnBrw 
         Height          =   435
         Index           =   3
         Left            =   1200
         TabIndex        =   7
         ToolTipText     =   "Stop"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":0060
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.chameleonButton btnBrw 
         Height          =   435
         Index           =   4
         Left            =   1560
         TabIndex        =   8
         ToolTipText     =   "Home"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":007C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.chameleonButton btnBrw 
         Height          =   435
         Index           =   5
         Left            =   1920
         TabIndex        =   9
         ToolTipText     =   "Search"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":0098
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin ProFile.chameleonButton btnBrw 
         CausesValidation=   0   'False
         Height          =   435
         Index           =   6
         Left            =   2280
         TabIndex        =   10
         ToolTipText     =   "Favorites"
         Top             =   0
         Width           =   435
         _ExtentX        =   767
         _ExtentY        =   767
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmBRW.frx":00B4
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblAddress 
         Caption         =   "&Address:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3165
         TabIndex        =   1
         Tag             =   "&Address:"
         Top             =   90
         Width           =   3075
      End
   End
End
Attribute VB_Name = "frmBRW"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim EventRunning As Boolean

Private Sub BRW_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    EventRunning = True
    cboAddress.Text = URL
    EventRunning = False
    Me.Caption = BRW.LocationName
End Sub

Private Sub BRW_NewWindow2(ppDisp As Object, Cancel As Boolean)
    On Error Resume Next
    If GetSet("Browser_AllowNewWindow", "1") = "1" Then
        Dim F As New frmBRW
        Set ppDisp = F.BRW.object
        F.Show
    Else
        Cancel = True
    End If
End Sub

Private Sub BRW_ProgressChange(ByVal Progress As Long, ByVal ProgressMax As Long)
    SProgress Progress, 0, ProgressMax
End Sub

Private Sub BRW_StatusTextChange(ByVal Text As String)
    SStatus Text
End Sub

Public Sub btnBrw_Click(Index As Integer)
    On Error Resume Next
    Select Case Index
        Case 0
            BRW.GoBack
        Case 1
            BRW.GoForward
        Case 2
            BRW.Refresh
        Case 3
            BRW.Stop
            Me.Caption = BRW.LocationName
        Case 4
            BRW.GoHome
        Case 5
            BRW.GoSearch
        Case 6
            frmMain.titBrowserFavorites_Click
    End Select
End Sub

Private Sub cboAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then
        Select Case Shift
            Case 0
                cboAddress_Click
            Case 2
                LoadFile "http://www." & cboAddress.Text & ".com"
        End Select
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    frmMain.titBrowser.Visible = True
    BRW.Silent = frmMain.titBrowserSilent.Checked
End Sub

Private Sub Form_Deactivate()
    On Error Resume Next
    frmMain.titBrowser.Visible = False
End Sub

Public Function OpenFavorites() As String
    On Error Resume Next
    frmBrwFav.Show 1
    OpenFavorites = FavStr
    FavStr = "" 'very important to clear this; otherwise cancel will be unusable
End Function

Private Sub Form_Load()
    On Error Resume Next
    Me.BRW.Silent = True
    BRW.GoHome
    Me.Icon = frmMain.Icon
    SkinForm Me
    SkinFormEx Me

    Dim i As Integer
    For i = 0 To 5 Step 1
        btnBrw(i + 1).Left = btnBrw(i).Left + btnBrw(i).Width + 15
    Next
    Form_Resize
End Sub

Private Sub BRW_DownloadComplete()
    On Error Resume Next
    EventRunning = True
    cboAddress.Text = BRW.LocationURL
    EventRunning = False
    Me.Caption = BRW.LocationName
End Sub

Public Function FavAddy(WhichFile As String) As String
    On Error Resume Next
        Dim A As String
        A = ReadINI("DEFAULT", "BASEURL", WhichFile)
        If Len(A) > 0 Then
            FavAddy = A
        Else
            A = ReadINI("InternetShortcut", "URL", WhichFile)
            If Len(A) > 0 Then
                FavAddy = A
            End If
        End If
End Function

Private Sub BRW_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
    On Error Resume Next
    Dim i As Integer
    Dim bFound As Boolean
    Me.Caption = BRW.LocationName
    EventRunning = True
    For i = 0 To cboAddress.ListCount - 1
        If cboAddress.List(i) = BRW.LocationURL Then
            bFound = True
            Exit For
        End If
    Next i
    If bFound Then
        cboAddress.RemoveItem i
    End If
    cboAddress.AddItem BRW.LocationURL, 0
    cboAddress.ListIndex = 0
    EventRunning = False
End Sub

Private Sub cboAddress_Click()
    On Error Resume Next
    If EventRunning Then Exit Sub
    LoadFile ParsedAddy(cboAddress.Text)
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    cboAddress.Width = Me.ScaleWidth - cboAddress.Left
    With BRW
        .Move 0, 480, Me.ScaleWidth, Me.ScaleHeight - 480 '480 being the top
    End With
End Sub

Public Function LoadFile(TheFN As String) As Long
    On Error Resume Next
    BRW.Navigate2 AddRecentItem(TheFN)
    Me.Caption = TrimFileNameLOL(TheFN)
    Me.Show
    Form_Resize
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Form_Deactivate
End Sub

