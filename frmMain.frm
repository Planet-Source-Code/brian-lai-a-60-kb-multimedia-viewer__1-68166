VERSION 5.00
Begin VB.MDIForm frmMain 
   Appearance      =   0  '¥­­±
   AutoShowChildren=   0   'False
   BackColor       =   &H8000000C&
   Caption         =   "ProFile"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   -210
   ClientWidth     =   10320
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   OLEDropMode     =   1  '¤â°Ê
   ScrollBars      =   0   'False
   StartUpPosition =   2  '¿Ã¹õ¤¤¥¡
   Begin VB.PictureBox picBrw 
      Align           =   3  '¹ï»ôªí³æ¥ª¤è
      BorderStyle     =   0  '¨S¦³®Ø½u
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6510
      Left            =   0
      ScaleHeight     =   6510
      ScaleWidth      =   2505
      TabIndex        =   0
      Top             =   0
      Width           =   2500
      Begin ProFile.chameleonButton btnFilerefresh 
         Height          =   375
         Left            =   0
         TabIndex        =   15
         ToolTipText     =   "Refresh file list"
         Top             =   0
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   661
         BTYPE           =   8
         TX              =   ""
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "·s²Ó©úÅé"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   16053492
         BCOLO           =   16053492
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   16777215
         MPTR            =   1
         MICON           =   "frmMain.frx":014A
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.PictureBox TrayIcon 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   255
         Left            =   1920
         Picture         =   "frmMain.frx":0166
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   14
         Top             =   2280
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.PictureBox Dragger 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   7350
         Index           =   1
         Left            =   2400
         MousePointer    =   9  'ªF-¦è¦V
         ScaleHeight     =   7350
         ScaleWidth      =   120
         TabIndex        =   5
         Top             =   0
         Width           =   120
      End
      Begin VB.DriveListBox Drive 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   360
         TabIndex        =   10
         Top             =   0
         Width           =   2175
      End
      Begin VB.PictureBox picFileTB 
         BorderStyle     =   0  '¨S¦³®Ø½u
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   0
         ScaleHeight     =   285.714
         ScaleMode       =   0  '¨Ï¥ÎªÌ¦Û­q
         ScaleWidth      =   2295
         TabIndex        =   4
         Top             =   6600
         Width           =   2295
         Begin VB.TextBox txtQuickFilter 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   285
            Left            =   0
            TabIndex        =   12
            Text            =   "Filter..."
            Top             =   0
            Width           =   1335
         End
         Begin VB.CommandButton btnSelType 
            Caption         =   "&Filter..."
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            TabIndex        =   11
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.PictureBox Dragger 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   120
         Index           =   0
         Left            =   0
         MousePointer    =   7  '¥_-«n¦V
         ScaleHeight     =   120
         ScaleWidth      =   2295
         TabIndex        =   3
         Top             =   2640
         Width           =   2295
      End
      Begin VB.FileListBox File 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3660
         Hidden          =   -1  'True
         Left            =   0
         System          =   -1  'True
         TabIndex        =   2
         Top             =   2760
         Width           =   2295
      End
      Begin VB.DirListBox Dir 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2250
         Left            =   0
         TabIndex        =   1
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.PictureBox picStatus 
      Align           =   2  '¹ï»ôªí³æ¤U¤è
      BorderStyle     =   0  '¨S¦³®Ø½u
      Height          =   300
      Left            =   0
      ScaleHeight     =   300
      ScaleWidth      =   10320
      TabIndex        =   6
      Top             =   6510
      Width           =   10320
      Begin VB.PictureBox picProgress 
         BorderStyle     =   0  '¨S¦³®Ø½u
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   1095
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
         Begin VB.Image iomgRandomImage 
            Height          =   225
            Left            =   0
            Top             =   0
            Width           =   225
         End
         Begin VB.Image ImgProgressVal 
            Height          =   225
            Left            =   240
            Picture         =   "frmMain.frx":02B0
            Stretch         =   -1  'True
            Top             =   0
            Width           =   375
         End
         Begin VB.Image ImgProgress 
            Height          =   225
            Left            =   0
            Picture         =   "frmMain.frx":02F6
            Stretch         =   -1  'True
            Top             =   0
            Width           =   15
         End
      End
      Begin VB.CommandButton btnSearch 
         Caption         =   "Sea&rch"
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
         Left            =   9480
         TabIndex        =   9
         Top             =   0
         Width           =   855
      End
      Begin VB.TextBox txtSearch 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   270
         Left            =   7560
         TabIndex        =   8
         Text            =   "Google..."
         Top             =   0
         Width           =   1935
      End
      Begin VB.Label lblStatus 
         AutoSize        =   -1  'True
         BackStyle       =   0  '³z©ú
         Caption         =   "Ready"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   30
         TabIndex        =   7
         Top             =   0
         UseMnemonic     =   0   'False
         Width           =   1425
      End
   End
   Begin VB.Menu titLE 
      Caption         =   "&PF"
      Begin VB.Menu titLENew 
         Caption         =   "New"
         Begin VB.Menu titLENewTextViewer 
            Caption         =   "Text Viewer"
         End
         Begin VB.Menu titLENewImageViewer 
            Caption         =   "Image Viewer"
         End
         Begin VB.Menu titLENewMediaPlayer 
            Caption         =   "Media Player"
         End
         Begin VB.Menu titLENewWebBrowser 
            Caption         =   "Web Browser"
            Shortcut        =   ^T
         End
      End
      Begin VB.Menu titPFOpenFile 
         Caption         =   "Open File..."
         Shortcut        =   ^O
      End
      Begin VB.Menu titLERecentFiles 
         Caption         =   "Recent Files"
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   8
         End
         Begin VB.Menu titLERecentFilesArray 
            Caption         =   ""
            Index           =   9
         End
      End
      Begin VB.Menu titLERecentFolders 
         Caption         =   "Recent Folders"
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   0
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   1
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   2
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   3
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   4
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   5
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   6
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   7
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   8
         End
         Begin VB.Menu titLERecentFoldersArray 
            Caption         =   ""
            Index           =   9
         End
      End
      Begin VB.Menu titClose 
         Caption         =   "Close"
         Begin VB.Menu titCloseThisWindow 
            Caption         =   "This Window"
            Shortcut        =   ^W
         End
         Begin VB.Menu titCloseAllWindows 
            Caption         =   "All Windows"
         End
      End
      Begin VB.Menu titS01 
         Caption         =   "-"
      End
      Begin VB.Menu titLEView 
         Caption         =   "View"
         Begin VB.Menu titLEViewRefreshFileList 
            Caption         =   "Refresh file list"
            Shortcut        =   ^{F5}
         End
         Begin VB.Menu titLEStatusBar 
            Caption         =   "Status Bar"
            Checked         =   -1  'True
         End
         Begin VB.Menu titLEViewAlwaysOnTop 
            Caption         =   "Always on top"
            Checked         =   -1  'True
         End
         Begin VB.Menu titLEViewMinToTray 
            Caption         =   "Minimize to tray"
         End
      End
      Begin VB.Menu titLETools 
         Caption         =   "Tools"
         Begin VB.Menu titLEToolsCalc 
            Caption         =   "EQ Calculator"
         End
         Begin VB.Menu titLEToolsFTP 
            Caption         =   "FTP Text Client"
         End
         Begin VB.Menu titLEToolsSlw 
            Caption         =   "Slideshow Maker"
         End
      End
      Begin VB.Menu titLESandbox 
         Caption         =   "Sandbox Mode"
         Checked         =   -1  'True
      End
      Begin VB.Menu titLEOptions 
         Caption         =   "Options..."
         Shortcut        =   ^R
      End
      Begin VB.Menu titLEAbout 
         Caption         =   "About PF"
      End
      Begin VB.Menu titLEExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu titFile 
      Caption         =   "&File"
      Begin VB.Menu titFileOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu titFileOpenAs 
         Caption         =   "Open as..."
      End
      Begin VB.Menu titFileOpenThisFolder 
         Caption         =   "Open this folder"
      End
      Begin VB.Menu titFileInfo 
         Caption         =   "Info..."
      End
      Begin VB.Menu titFileShell 
         Caption         =   "Shell"
         Begin VB.Menu titFileShellOpen 
            Caption         =   "Open"
         End
         Begin VB.Menu titFileShellEdit 
            Caption         =   "Edit"
         End
      End
      Begin VB.Menu titFileCopyTo 
         Caption         =   "Copy to..."
      End
      Begin VB.Menu titFileMoveTo 
         Caption         =   "Move to..."
      End
      Begin VB.Menu titFileRename 
         Caption         =   "Rename..."
      End
      Begin VB.Menu titFileMoveToRecycleBin 
         Caption         =   "Move to recycle bin"
      End
      Begin VB.Menu titFileDelete 
         Caption         =   "Delete"
      End
   End
   Begin VB.Menu titText 
      Caption         =   "&Text"
      Begin VB.Menu titTextFile 
         Caption         =   "File"
         Begin VB.Menu titTextFileOpen 
            Caption         =   "Open File..."
         End
         Begin VB.Menu titTextFileOpenURL 
            Caption         =   "Open URL..."
         End
         Begin VB.Menu titTextFileSave 
            Caption         =   "Save"
         End
         Begin VB.Menu titTextFileSaveAs 
            Caption         =   "Save as..."
         End
      End
      Begin VB.Menu titTextEdit 
         Caption         =   "Edit"
         Begin VB.Menu titTextEditCut 
            Caption         =   "Cut"
         End
         Begin VB.Menu titTextEditCopy 
            Caption         =   "Copy"
         End
         Begin VB.Menu titTextEditPaste 
            Caption         =   "Paste"
         End
         Begin VB.Menu titS02 
            Caption         =   "-"
         End
         Begin VB.Menu titTextEditSelectAll 
            Caption         =   "Select All"
         End
      End
      Begin VB.Menu titTextView 
         Caption         =   "View"
         Begin VB.Menu titTextViewFont 
            Caption         =   "Font..."
         End
         Begin VB.Menu titTextEditSelText 
            Caption         =   "Selected text..."
            Begin VB.Menu titTextViewSelTextOpen 
               Caption         =   "Open"
            End
            Begin VB.Menu titTextViewSelTextOpenAsWeb 
               Caption         =   "Open as web page"
               Shortcut        =   ^{F3}
            End
            Begin VB.Menu titTextViewSelTextOpenAsImage 
               Caption         =   "Open as image"
            End
            Begin VB.Menu titTextViewSelTextOpenAsMedia 
               Caption         =   "Open as Media"
            End
         End
         Begin VB.Menu titS05 
            Caption         =   "-"
         End
         Begin VB.Menu titTextViewRunThisCode 
            Caption         =   "Run this code"
            Shortcut        =   {F5}
         End
         Begin VB.Menu titTextViewRunSelection 
            Caption         =   "Run Selection"
         End
      End
   End
   Begin VB.Menu titMedia 
      Caption         =   "&Media"
      Begin VB.Menu titMediaFile 
         Caption         =   "File"
         Begin VB.Menu titMediaFileOpen 
            Caption         =   "Open File..."
         End
         Begin VB.Menu titMediaFileOpenURL 
            Caption         =   "Open from URL..."
         End
         Begin VB.Menu titMediaFileOpenClipboardFileName 
            Caption         =   "Open clipboard file name"
         End
      End
      Begin VB.Menu titMediaStretchVideo 
         Caption         =   "Stretch Video"
         Checked         =   -1  'True
      End
      Begin VB.Menu titMediaControls 
         Caption         =   "Controls"
         Checked         =   -1  'True
      End
      Begin VB.Menu titMediaSyncPSM 
         Caption         =   "Sync with MSN Messenger"
         Checked         =   -1  'True
      End
      Begin VB.Menu titMediaPlaySongOnStartup 
         Caption         =   "Play song on start up"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu titImage 
      Caption         =   "&Image"
      Begin VB.Menu titImageFile 
         Caption         =   "File"
         Begin VB.Menu titImageFileOpen 
            Caption         =   "Open File..."
         End
         Begin VB.Menu titImageFileOpenURL 
            Caption         =   "Open URL..."
         End
      End
      Begin VB.Menu titImageBorder 
         Caption         =   "Border"
         Checked         =   -1  'True
      End
      Begin VB.Menu titImageStretch 
         Caption         =   "Stretch"
         Checked         =   -1  'True
      End
      Begin VB.Menu titImageCheckers 
         Caption         =   "Checkers on background"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu titBrowser 
      Caption         =   "&Browser"
      Begin VB.Menu titBrowserFile 
         Caption         =   "File"
         Begin VB.Menu titBrowserFileOpen 
            Caption         =   "Open File..."
         End
         Begin VB.Menu titBrowserFileOpenURL 
            Caption         =   "Open URL..."
         End
      End
      Begin VB.Menu titBrowserZoom 
         Caption         =   "Zoom"
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "10%"
            Index           =   0
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "20%"
            Index           =   1
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "30%"
            Index           =   2
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "40%"
            Index           =   3
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "50%"
            Index           =   4
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "60%"
            Index           =   5
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "70%"
            Index           =   6
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "80%"
            Index           =   7
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "90%"
            Index           =   8
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "100%"
            Index           =   9
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "110%"
            Index           =   10
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "120%"
            Index           =   11
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "130%"
            Index           =   12
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "140%"
            Index           =   13
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "150%"
            Index           =   14
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "160%"
            Index           =   15
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "170%"
            Index           =   16
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "180%"
            Index           =   17
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "190%"
            Index           =   18
         End
         Begin VB.Menu titBrowserZoomArray 
            Caption         =   "200%"
            Index           =   19
         End
         Begin VB.Menu titS04 
            Caption         =   "-"
         End
         Begin VB.Menu titBrowserZoomFull 
            Caption         =   "Full Screen"
         End
      End
      Begin VB.Menu titBrowserFavorites 
         Caption         =   "Favorites"
         Shortcut        =   ^{F2}
      End
      Begin VB.Menu titBrowserSilent 
         Caption         =   "Silent"
         Checked         =   -1  'True
      End
      Begin VB.Menu titBrowserAllowNewWindow 
         Caption         =   "Allow New Window"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu titWindows 
      Caption         =   "&Windows"
      WindowList      =   -1  'True
      Begin VB.Menu titWindowsMaxAll 
         Caption         =   "Maximize"
         Shortcut        =   {F1}
      End
      Begin VB.Menu titWindowsMin 
         Caption         =   "Restore"
      End
      Begin VB.Menu titWindowsTileAll 
         Caption         =   "Tile"
      End
      Begin VB.Menu titWindowsTileH 
         Caption         =   "Tile Horizontally"
         Shortcut        =   {F2}
      End
      Begin VB.Menu titWindowsTileV 
         Caption         =   "Tile Vertically"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OldX As Long, OldY As Long
Const HarryIsDumb As Boolean = True
Const BarWidth As Long = 2500
'Do Not Show again UBound: Use 10

Private Sub btnFilerefresh_Click()
    On Error Resume Next
    titLEViewRefreshFileList_Click
End Sub

Private Sub btnSearch_Click()
    On Error Resume Next
    Dim A As String
    A = txtSearch.Text
    If A = GetSet("Search_Provider_Name", DefaultSearchAgent) & "..." Then Exit Sub
    A = Replace(txtSearch.Tag, "%s", A)
    If Len(A) = 0 Then Exit Sub
    Dim B As New frmBRW
    B.BRW.Navigate A
    B.Show
End Sub

Private Sub btnSelType_Click()
    On Error Resume Next
    frmShowFileFilter.Show 1
End Sub

Private Sub Dir_Change()
    On Error Resume Next
    File.Path = Dir.Path
    Drive.Drive = Dir.Path
    SaveSet "Recent_Path", Dir.Path
End Sub

Private Sub Dragger_DblClick(Index As Integer)
    On Error Resume Next
    If Index = 1 Then
        If picBrw.Width = BarWidth Then
            picBrw.Width = Dragger(1).Width
        Else
            picBrw.Width = BarWidth
        End If
        picBrw_Resize
    End If
End Sub

Private Sub Dragger_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    OldX = X: OldY = Y
    Dragger(Index).BackColor = RGB(255, 0, 0)
End Sub

Private Sub Dragger_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim A As Long
    Select Case Button
        Case 1
            Select Case Index
                Case 0
                    A = Dragger(0).Top - OldY + Y
                    A = A - 60 - A Mod 240 'to disable integral height thingy
                    If A > Me.ScaleHeight - 300 Then A = Me.ScaleHeight - Dragger(0).Height
                    If A < Drive.Height Then A = Drive.Height
                    Dragger(0).Move 0, A
                Case 1
                    A = picBrw.Width - OldX + X
                    If A < Dragger(1).Width + 300 Then A = Dragger(1).Width   'Redraw
                    If A > Me.Width - 300 Then A = Me.Width - Dragger(1).Width
                    If A < BarWidth + 300 And A > BarWidth - 300 Then A = BarWidth
                    If A < Me.Width / 2 + 300 And A > Me.Width / 2 - 300 Then A = Me.Width / 2
                    Drive.Visible = (A > 300)
                    picBrw.Width = A
                    X = OldX
                    A = Null
            End Select
    End Select
End Sub

Private Sub Dragger_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next 'handles refresh
    picBrw_Resize
    Select Case Index
        Case 0
            SaveSet "Dragger0_Top", Dragger(0).Top
        Case 1
            SaveSet "picBrw_Width", picBrw.Width
    End Select
    Dragger(Index).BackColor = picBrw.BackColor '&H8000000F
End Sub

Private Sub Drive_Change()
    On Error Resume Next
    Dir.Path = Me.Drive
End Sub

Private Sub File_DblClick()
    On Error Resume Next
    AddRecentFolder File.Path
    DecideOnType FindPath(File.Path, File.filename), File.filename
End Sub

Public Function DecideOnType(eFilePathPlusName As String, eFileName As String, Optional IgnoreRMBFileExtFlag As Boolean)
    On Error Resume Next
    Dim A As String, G As String
    Dim E As Long
    G = Right$(eFileName, Len(eFileName) - InStrRev(eFileName, "."))
    E = OpenFileDlg.AsType(G, , IgnoreRMBFileExtFlag)
    Select Case E
        Case 0 'text
            Dim B As New frmTXT
            B.LoadFile eFilePathPlusName
        Case 1 'media
            Dim C As New frmWMP
            C.LoadFile eFilePathPlusName
        Case 2 'image
            Dim D As New frmIMG
            D.LoadFile eFilePathPlusName
        Case 3 'web
            Dim F As New frmBRW
            F.LoadFile eFilePathPlusName
        Case 4 'bookmark
            F.LoadFile F.FavAddy(eFilePathPlusName)
        Case 5 'default
            Call ShellExecute(Me.hwnd, "open", eFilePathPlusName, "", File.Path, 1)
        Case 6 'the "I HAVE NO IDEA" option
            F.LoadFile "filext.com/detaillist.php?extdetail=" & G
        Case 99 'fails
            Exit Function
    End Select
    If E < 3 Then SStatus App.ProductName & " opened " & eFileName
End Function

Private Sub File_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button = 2 Then
        File_MouseMove Button, Shift, X, Y
    End If
End Sub

Private Sub File_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    Dim Ix As Long
    Dim Mx As Long, My As Long
    SStatus "Double click to open. Right click for more options."
    If Button = 2 Then
        Mx = CLng(X / Screen.TwipsPerPixelX)
        My = CLng(Y / Screen.TwipsPerPixelY)
        Ix = SendMessage(File.hwnd, LB_ITEMFROMPOINT, 0, ByVal ((My * 65536) + Mx))
        If Ix < File.ListCount Then
            File.Selected(Ix) = True
            PopupMenu titFile, , Mx * Screen.TwipsPerPixelX + File.Left, My * Screen.TwipsPerPixelY + File.Top, titFileOpen
        End If
    End If
End Sub

Private Sub MDIForm_Initialize()
    On Error Resume Next
    If App.PrevInstance Then
        MsgBox "Another instance is running - please close it before launching another one.", vbExclamation
        End
    End If
End Sub

Private Sub MDIForm_Load()
    On Error Resume Next
    OnTop Me.hwnd, (GetSet("MDIForm_OnTop", "0") = "1")
    Me.WindowState = Val(GetSet("MDIForm_WinMode"))
    Me.Width = Val(GetSet("MDIForm_Width", Str(Me.Width)))
    Me.Height = Val(GetSet("MDIForm_Height", Str(Me.Height)))
    Dragger(0).Top = Val(GetSet("Dragger0_Top", Str(Dragger(0).Top)))
    picBrw.Width = Val(GetSet("picBrw_Width", Str(picBrw.Width)))
    txtSearch.Text = GetSet("Search_Provider_Name", DefaultSearchAgent) & "..."
    txtQuickFilter.Text = DisplayFilter
    If GetSet("Search_Fuzzy", "1") = "1" Then
        File.Pattern = "*" & GetSet("File_Pattern") & "*"
    Else
        File.Pattern = GetSet("File_Pattern", "*.*")
    End If
    Dir.Path = GetSet("Recent_Path", Dir.Path)
    picStatus.Visible = (GetSet("Status_Bar", "1") = "1")
    'Mod32BitIcon.SetIcon Me.hwnd, "AAA"
    Call LoadRecents
    Call LoadRecentFolders
    Call OptionalMenus(False) 'Yes I do this on purpose - allows design-time menu access
    Call LoadCheckBoxes
        
    SkinForm Me
    SkinFormEx Me

    picBrw_Resize
    Call LoadSearchProvider
    'startup loader
    Select Case GetSet("OpenOnStart")
        Case "1"
            titLENewTextViewer_Click
        Case "2"
            titLENewImageViewer_Click
        Case "3"
            titLENewMediaPlayer_Click
            If GetSet("Media_StartPlay", "1") = "1" Then
                AF.LoadFile GetSet("Media_Last")
            End If
        Case "4"
            titLENewWebBrowser_Click
    End Select
    SkinForm Me
    DSA 2
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub LoadCheckBoxes()
    On Error Resume Next
    titMediaControls.Checked = (GetSet("Media_Controls", "1") = "1")
    titMediaPlaySongOnStartup.Checked = (GetSet("Media_StartPlay", "0") = "1")
    titMediaStretchVideo.Checked = (GetSet("Media_Stretch", "1") = "1")
    titMediaSyncPSM.Checked = (GetSet("Sync_PSM", "1") = "1")
    titImageBorder.Checked = (GetSet("Image_Border", "1") = "1")
    titImageCheckers.Checked = (GetSet("Image_Checkers", "1") = "1")
    titImageStretch.Checked = (Val(GetSet("Image_Stretch", "2")) >= 1)
    titBrowserAllowNewWindow.Checked = (GetSet("Browser_AllowNewWindow", "1") = "1")
    titBrowserSilent.Checked = (GetSet("Browser_Silent", "1") = "1")
    titLESandbox.Checked = (GetSet("Sandbox") = "1")
    titLEStatusBar.Checked = (GetSet("Status_Bar") = "1")
    titLEViewAlwaysOnTop.Checked = (GetSet("MDIForm_OnTop", "0") = "1")
End Sub

Private Sub OptionalMenus(Optional TrueOrFalse As Boolean = True)
    On Error Resume Next
    titFile.Visible = TrueOrFalse
    titText.Visible = TrueOrFalse
    titMedia.Visible = TrueOrFalse
    titImage.Visible = TrueOrFalse
    titBrowser.Visible = TrueOrFalse
End Sub

Private Sub MDIForm_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'On Error Resume Next
    Dim i As Integer
    For i = 1 To Data.Files.Count Step 1
        DecideOnType Data.Files.item(i), TrimFileNameLOL(Data.Files.item(i))
    Next
End Sub

Private Sub MDIForm_Resize()
    On Error Resume Next
    SaveSet "MDIForm_Width", Me.Width
    SaveSet "MDIForm_Height", Me.Height
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    On Error Resume Next
    DeleteIcon TrayIcon
    File.Pattern = Replace(File.Pattern, "**", "*")
    File.Pattern = Replace(File.Pattern, ".*.", ".")
    SaveSet "File_Pattern", File.Pattern
    SaveSet "MDIForm_WinMode", Me.WindowState
    End
End Sub

Private Sub picBrw_Resize()
    On Error Resume Next
    Dragger(0).Width = picBrw.Width - Dragger(1).Width
    Dragger(1).Move picBrw.Width - Dragger(1).Width, 0, Dragger(1).Width, picBrw.Height
    Dir.Move 0, Drive.Height, picBrw.Width - Dragger(1).Width, Dragger(0).Top - Drive.Height
    Dim A As Long: A = Dragger(0).Top + Dragger(0).Height
    File.Move 0, A, picBrw.Width - Dragger(1).Width, picBrw.Height - A - picFileTB.Height - 30
    picFileTB.Move 0, File.Top + File.Height + 15, picBrw.Width - Dragger(1).Width
    btnFilerefresh.Height = Drive.Height
    Drive.Move btnFilerefresh.Width, Drive.Top, picFileTB.Width - Drive.Left ', btnFilerefresh.Height
    btnFilerefresh.Visible = (picBrw.Width > 300)
End Sub

Private Sub picFileTB_Resize()
    On Error Resume Next
    btnSelType.Move picFileTB.Width - btnSelType.Width, 0, btnSelType.Width, picFileTB.Height - 15
    txtQuickFilter.Move 0, txtQuickFilter.Top, picFileTB.Width - btnSelType.Width - 30
End Sub

Private Sub picStatus_Resize()
    On Error Resume Next
    btnSearch.Move picStatus.Width - btnSearch.Width, 15, btnSearch.Width, picStatus.Height - 15
    txtSearch.Move btnSearch.Left - txtSearch.Width - 15, 15, txtSearch.Width, picStatus.Height - 30
    'lblStatus.Move 0, 15, txtSearch.Left, 225 '225 as in, to prevent showing of second line
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub titBrowserAllowNewWindow_Click()
    On Error Resume Next
    With titBrowserAllowNewWindow
        .Checked = Not .Checked
        SaveSet "Browser_AllowNewWindow", IIf(.Checked, "1", "0")
    End With
End Sub

Public Sub titBrowserFavorites_Click()
    Dim A As String
    A = AF.OpenFavorites() 'allows oppurtunity to detect len
    If Len(A) > 0 Then AF.LoadFile A
End Sub

Private Sub titBrowserFileOpen_Click()
    titTextFileOpen_Click
End Sub

Private Sub titBrowserFileOpenURL_Click()
    On Error Resume Next
    Dim A As String
    A = InputBox("Enter URL or Location:", , Clipboard.GetText)
    If Len(A) > 0 Then AF.LoadFile A
End Sub

Private Sub titBrowserSilent_Click()
    On Error Resume Next
    With titBrowserSilent
        .Checked = Not .Checked
        SaveSet "Browser_Silent", IIf(.Checked, "1", "0")
        AF.BRW.Silent = .Checked
    End With
End Sub

Private Sub titBrowserZoomArray_Click(Index As Integer)
    On Error Resume Next
    AF.BRW.Document.body.Style.Zoom = titBrowserZoomArray(Index).Caption
End Sub

Private Sub titBrowserZoomFull_Click()
    On Error Resume Next
    DSA 10
    Shell FindPath(App.Path, "TEExt.exe ") & AF.BRW.LocationURL, vbNormalFocus
End Sub

Private Sub titCloseAllWindows_Click()
    titLEExit_Click
End Sub

Private Sub titCloseThisWindow_Click()
    On Error Resume Next
    Unload AF
End Sub

Private Sub titFileCopyTo_Click()
    On Error Resume Next
    File.Tag = "Copy To..."
    JustDoIt1
    File.Refresh
End Sub

Private Function JustDoIt1() As Boolean
    On Error Resume Next
    Dim A As String
    A = BrowseForFolder(Me.hwnd, File.Tag)
    If Len(A) > 0 Then
        FileCopy FindPath(File.Path, File.filename), FindPath(A, File.filename)
        JustDoIt1 = True
    End If
End Function

Private Sub titFileDelete_Click()
    On Error Resume Next
    If MsgBox("Are you sure you want to delete " & File.filename & " permanently?", vbYesNo + vbQuestion) = vbYes Then
        Kill FindPath(File.Path, File.filename)
        File.Refresh
        SStatus File.filename & " deleted permanently"
    End If
End Sub

Private Sub titFileInfo_Click()
    On Error Resume Next
    Dim A As String, B As String
    A = FindPath(File.Path, File.filename)
    If GetAttrib(A, vbArchive) = True Then B = B & "Archived"
    If GetAttrib(A, vbCompressed) = True Then B = B & ", Compressed"
    If GetAttrib(A, vbDirectory) = True Then B = B & ", Directory"
    If GetAttrib(A, vbHidden) = True Then B = B & ", Hidden"
    If GetAttrib(A, vbNormal) = True Then B = B & ", Normal"
    If GetAttrib(A, vbReadOnly) = True Then B = B & ", Read Only"
    If GetAttrib(A, vbTemporary) = True Then B = B & ", Temporary"
    If GetAttrib(A, vbVolume) = True Then B = B & ", Volume"
    If Left$(B, 2) = ", " Then B = Mid$(B, 3)
    MyMsgBox A & vbCrLf & vbCrLf & _
                    B & vbCrLf & vbCrLf & _
                    Round(Val(FileLen(A)) / 1024 / 1024, 2) & " MB", 8, "File Info", True
End Sub

Private Sub titFileMoveTo_Click()
    On Error Resume Next
    File.Tag = "Move To..."
    If JustDoIt1 = True Then Kill FindPath(File.Path, File.filename) 'delete teh file
    File.Refresh
End Sub

Private Sub titFileMoveToRecycleBin_Click()
    On Error Resume Next
    Dim typOperation As SHFILEOPSTRUCT
    With typOperation
            .wFunc = &H3
            .pFrom = FindPath(File.Path, File.filename)
            .fFlags = &H40
        End With
        SHFileOperation typOperation
    File.Refresh
End Sub

Private Sub titFileOpen_Click()
    On Error Resume Next
    File_DblClick
End Sub

Private Sub titFileOpenAs_Click()
    On Error Resume Next
    AddRecentFolder File.Path
    DecideOnType FindPath(File.Path, File.filename), File.filename, True
End Sub

Private Sub titFileOpenThisFolder_Click()
    On Error Resume Next
    Shell "explorer " & File.Path
End Sub

Private Sub titFileRename_Click()
    On Error Resume Next
    Dim A As String, B As String
    With File
        B = .filename
        A = InputBox("Enter new file name:", "Rename", B)
        If Len(A) > 0 Then
            Name FindPath(.Path, B) As FindPath(.Path, A) 'rename kewl!!1@
            .Refresh
        End If
    End With
End Sub

Private Sub titFileShellEdit_Click()
    On Error Resume Next
    Call ShellExecute(Me.hwnd, "edit", FindPath(File.Path, File.filename), "", File.Path, 1)
End Sub

Private Sub titFileShellOpen_Click()
    On Error Resume Next
    Call ShellExecute(Me.hwnd, "open", FindPath(File.Path, File.filename), "", File.Path, 1)
End Sub

Private Sub titImageBorder_Click()
    With titImageBorder
        .Checked = Not .Checked
        SaveSet "Image_Border", IIf(.Checked, "1", "0")
        AF.IMG.BorderStyle = IIf(.Checked, 1, 0)
    End With
End Sub

Private Sub titImageCheckers_Click()
    On Error Resume Next
    With titImageCheckers
        .Checked = Not .Checked
        SaveSet "Image_Checkers", IIf(.Checked, "1", "0")
        AF.imgBG.Visible = .Checked
        AF.Form_Resize
    End With
End Sub

Private Sub titImageFileOpen_Click()
    titTextFileOpen_Click
End Sub

Private Sub titImageFileOpenURL_Click()
    On Error Resume Next
    Dim A As String
    DSA 4
    A = InputBox("Enter URL or Location:", , Clipboard.GetText)
    If Len(A) > 0 Then AF.LoadFile DownloadFile(A)
End Sub

Private Sub titImageStretch_Click()
    On Error Resume Next
    With titImageStretch
        .Checked = Not .Checked
        SaveSet "Image_Stretch", IIf(.Checked, "2", "0")
        AF.DoStretch Val(GetSet("Image_Stretch", "2")) 'launch 1 if you want...
        AF.Form_Resize
    End With
End Sub

Private Sub titLEAbout_Click()
    On Error Resume Next
    frmPrefs.GoToTab 0
    frmPrefs.Show 1
End Sub

Private Sub titLEExit_Click()
    On Error Resume Next
    End
End Sub

Private Sub titLERecentFilesArray_Click(Index As Integer)
    On Error Resume Next
    DecideOnType titLERecentFilesArray(Index).Tag, titLERecentFilesArray(Index).Caption
End Sub

Private Sub titLENewImageViewer_Click()
    On Error Resume Next
    Dim B As New frmIMG
    B.Show
End Sub

Private Sub titLENewMediaPlayer_Click()
    On Error Resume Next
    Dim B As New frmWMP
    B.Show
End Sub

Private Sub titLENewTextViewer_Click()
    On Error Resume Next
    Dim B As New frmTXT
    B.Show
End Sub

Private Sub titLENewWebBrowser_Click()
    On Error Resume Next
    Dim B As New frmBRW
    B.Show
End Sub

Private Sub titLEOptions_Click()
    On Error Resume Next
    frmPrefs.Show 1
End Sub

Private Sub titLERecentFoldersArray_Click(Index As Integer)
    On Error Resume Next
    Dim Emerson As String
    Emerson = titLERecentFoldersArray(Index).Tag
    Drive.Drive = Left$(Emerson, 1)
    Dir.Path = Emerson
    File.Path = Emerson
End Sub

Private Sub titLESandbox_Click()
    On Error Resume Next
    DSA 5
    If GetSet("Sandbox") = "1" Then
        DSA 1
    End If
    With titLESandbox
        .Checked = Not .Checked
'        WriteINI UserName, "Sandbox", IIf(.Checked, "1", "0"), SettingsFile
        SaveSet "Sandbox", IIf(.Checked, "1", "0")
    End With
End Sub

Private Sub titLEStatusBar_Click()
    On Error Resume Next
    With titLEStatusBar
        .Checked = Not .Checked
        SaveSet "Status_Bar", IIf(.Checked, "1", "0")
        picStatus.Visible = .Checked
    End With
End Sub

Private Sub titLEToolsCalc_Click()
    On Error Resume Next
    DSA 10
    Shell FindPath(App.Path, "TEExt.exe tetracal"), vbNormalFocus
End Sub

Private Sub titLEToolsFTP_Click()
    On Error Resume Next
    DSA 10
    Shell FindPath(App.Path, "TEExt.exe TetraFTP"), vbNormalFocus
End Sub

Private Sub titLEToolsSlw_Click()
    On Error Resume Next
    DSA 10
    Shell FindPath(App.Path, "TEExt.exe tetraslw ") & File.Path, vbNormalFocus
End Sub

Private Sub titLEViewAlwaysOnTop_Click()
    On Error Resume Next
    With titLEViewAlwaysOnTop
        .Checked = Not .Checked
        SaveSet "MDIForm_OnTop", IIf(.Checked, "1", "0")
        OnTop Me.hwnd, .Checked
    End With
End Sub

Private Sub titLEViewMinToTray_Click()
'    On Error Resume Next
    NoSysIcon False
End Sub

Private Sub titLEViewRefreshFileList_Click()
    On Error Resume Next
    Dim A As String, B As String
    A = Left$(App.Path, 3) 'like, C:\ or D:\
    B = File.Path
    Drive.Drive = A
    Dir.Path = A
    File.Path = A
    Drive.Drive = Left$(B, 3) 'like, C:\ or D:\
    Dir.Path = B
    File.Path = B
End Sub

Private Sub titMediaControls_Click()
    On Error Resume Next
    With titMediaControls
        .Checked = Not .Checked
        SaveSet "Media_Controls", IIf(.Checked, "1", "0")
        Me.ActiveForm.WMP.uiMode = IIf(GetSet("Media_Controls", "1") = "1", "full", "none")
    End With
End Sub

Private Sub titMediaFileOpen_Click()
    On Error Resume Next
    titTextFileOpen_Click
End Sub

Private Sub titMediaFileOpenClipboardFileName_Click()
    On Error Resume Next
    AF.LoadFile Clipboard.GetText
End Sub

Private Sub titMediaFileOpenURL_Click()
    On Error Resume Next
    Dim A As String
    A = InputBox("Enter URL or location here:", , AF.WMP.URL)
    If Len(A) > 0 Then
        AF.LoadFile A
    End If
End Sub

Private Sub titMediaPlaySongOnStartup_Click()
    On Error Resume Next
    With titMediaPlaySongOnStartup
        .Checked = Not .Checked
        If .Checked Then
            DSA 6
            SaveSet "OpenOnStart", "3" 'so a media pops up
        End If
        SaveSet "Media_StartPlay", IIf(.Checked, "1", "0")
    End With
End Sub

Private Sub titMediaStretchVideo_Click()
    On Error Resume Next
    With titMediaStretchVideo
        .Checked = Not .Checked
        SaveSet "Media_Stretch", IIf(.Checked, "1", "0")
        AF.WMP.stretchToFit = GetSet("Media_Stretch", "1")
    End With
End Sub

Private Sub titMediaSyncPSM_Click()
    On Error Resume Next
    With titMediaSyncPSM
        .Checked = Not .Checked
        If .Checked Then
            DSA 7
        End If
        SaveSet "Sync_PSM", IIf(.Checked, "1", "0")
    End With
End Sub

Private Sub titPFOpenFile_Click()
    On Error Resume Next
    Dim Response As VbMsgBoxResult
    With cmndlg
        .filefilter = "All Files|*.*|" & _
                        "Text Files|*.txt;*.dat;*.ini;*.sys;*.htm;*.html;*.xml|" & _
                        "Media Files|*.wav;*.mp1;*.mp2;*.mp3;*.mpg;*.mpeg;*.m4a;*.wma;*.wmv;*.mid;*.aiff;*.dat|" & _
                        "Image Files|*.jpg;*.jpeg;*.jpe;*.gif;*.bmp"
        OpenFile
        If Len(.filename) = 0 Then Exit Sub
        DecideOnType .filename, TrimFileNameLOL(.filename)
    End With
End Sub

Private Sub titTextEditCopy_Click()
    On Error Resume Next
    If LCase$(Me.ActiveForm.Name) = "frmtxt" Then 'identify form
        With AF.txtBox
            Clipboard.SetText .SelText
        End With
    End If
End Sub

Private Sub titTextEditCut_Click()
    On Error Resume Next
    If LCase$(AF.Name) = "frmtxt" Then 'identify form
        With AF.txtBox
            Clipboard.SetText .SelText
            .SelText = ""
        End With
    End If
End Sub

Private Sub titTextEditPaste_Click()
    On Error Resume Next
    If LCase$(AF.Name) = "frmtxt" Then 'identify form
        With AF.txtBox
            .SelText = Clipboard.GetText
        End With
    End If
End Sub

Private Sub titTextEditSelectAll_Click()
    On Error Resume Next
    With AF.txtBox
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub titTextFileOpen_Click()
    On Error Resume Next
    Dim Response As VbMsgBoxResult
    With cmndlg
        .filefilter = "any file (*.*)|*.*"
        OpenFile
        If Len(.filename) = 0 Then Exit Sub
        AF.LoadFile .filename
    End With
End Sub

Private Sub titTextFileOpenURL_Click()
    On Error Resume Next
    titImageFileOpenURL_Click
End Sub

Public Sub titTextFileSave_Click()
    On Error Resume Next
    If Len(AF.Tag) = 0 Then
        titTextFileSaveAs_Click 'if theres no tag... save as!
        Exit Sub
    End If
    TXTFileSave AF.txtBox.Text, AF.Tag
    If Right$(AF.Caption, 1) = "*" Then AF.Caption = Left$(AF.Caption, Len(AF.Caption) - 1)
End Sub

Private Sub titTextFileSaveAs_Click()
    With cmndlg
        .filefilter = "All files (*.*)|*.*"
        .flags = 5 Or 2
        SaveFile
        If Len(.filename) = 0 Then Exit Sub
        AF.Tag = .filename
    End With
    titTextFileSave_Click
    AF.Caption = TrimFileNameLOL(AF.Tag)
End Sub

Private Sub titTextViewFont_Click()
    On Error Resume Next
    AF.ChangeFont
End Sub

Private Sub titTextViewRunSelection_Click()
    On Error Resume Next
    Dim A As String
    DSA 3
    A = Replace(AF.txtBox.SelText, vbCrLf, "|") 'the SEL
    Shell FindPath(App.Path, "ESE.exe") & " " & A
End Sub

Private Sub titTextViewRunThisCode_Click()
    On Error Resume Next
    Dim A As String
    DSA 3
    A = Replace(AF.txtBox.Text, vbCrLf, "|")
    Shell FindPath(App.Path, "ESE.exe") & " " & A
End Sub

Private Sub titTextViewSelTextOpen_Click()
    On Error Resume Next
    DecideOnType AF.txtBox.SelText, TrimFileNameLOL(AF.txtBox.SelText)
End Sub

Private Sub titTextViewSelTextOpenAsImage_Click()
    On Error Resume Next
    Dim A As New frmIMG
    A.LoadFile AF.txtBox.SelText
End Sub

Private Sub titTextViewSelTextOpenAsMedia_Click()
    On Error Resume Next
    Dim A As New frmWMP
    A.LoadFile AF.txtBox.SelText
End Sub

Private Sub titTextViewSelTextOpenAsWeb_Click()
    On Error Resume Next
    Dim A As New frmBRW
    A.LoadFile AF.txtBox.SelText
End Sub

Private Sub titWindowsMaxAll_Click()
    On Error Resume Next
    AF.WindowState = 2
End Sub

Private Sub titWindowsMin_Click()
    On Error Resume Next
    AF.WindowState = 0
End Sub

Private Sub titWindowsTileAll_Click()
    On Error Resume Next
    Me.Arrange 0
End Sub

Private Sub titWindowsTileH_Click()
    On Error Resume Next
    Me.Arrange 1
End Sub

Private Sub titWindowsTileV_Click()
    On Error Resume Next
    Me.Arrange 2
End Sub

Private Sub Trayicon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Msg As Long
    Msg = (X And &HFF) * &H100
    Select Case Msg
        Case 0 'mouse moves
        Case &HF00  'left mouse button down
        Case &H1E00 'left mouse button up
        Case &H3C00  'right mouse button down
        Case &H2D00 'left mouse button double click
        NoSysIcon True    'Show App on double clicking Mouse's Left Button
        Case &H4B00 'right mouse button up
        Case &H5A00 'right mouse button double click
    End Select
End Sub

Private Sub txtQuickFilter_Change()
    On Error Resume Next
    With txtQuickFilter
        If .Text = GetSet("Filtre_String", DefaultFilterString) & "..." Or .Text = "" Then
            File.Pattern = "*.*"
        Else
            If GetSet("Search_Fuzzy", "1") = "1" Then
                File.Pattern = "*" & .Text & "*"
            Else
                File.Pattern = .Text
            End If
        End If
    End With
End Sub

Private Sub txtQuickFilter_GotFocus()
    On Error Resume Next
    With txtQuickFilter
        .ForeColor = RGB(0, 0, 0)
        If .Text = GetSet("Filtre_String", DefaultFilterString) & "..." Then
            .Text = ""
        End If
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtQuickFilter_LostFocus()
    On Error Resume Next
    With txtQuickFilter
        .ForeColor = RGB(127, 127, 127)
        If .Text = "" Then
            .Text = GetSet("Filtre_String", DefaultFilterString) & "..."
        End If
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSearch_GotFocus()
    On Error Resume Next
    With txtSearch
        .ForeColor = RGB(0, 0, 0)
        If .Text = GetSet("Search_Provider_Name", DefaultSearchAgent) & "..." Then
            .Text = ""
        End If
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = vbKeyReturn Then btnSearch_Click
End Sub

Private Sub txtSearch_LostFocus()
    On Error Resume Next
    With txtSearch
        .ForeColor = RGB(127, 127, 127)
        If .Text = "" Then
            .Text = GetSet("Search_Provider_Name", DefaultSearchAgent) & "..."
        End If
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub
