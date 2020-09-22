Attribute VB_Name = "ModSysTray"
'######################################################################
'System Tray Declarations Starts
'######################################################################

Public INTRAY As Boolean 'Boolean to detect App Status[Max or Min]

' Declare Tray Icon
Type NOTIFYICONDATA
     cbSize As Long
     hwnd As Long
     uID As Long
     uFlags As Long
     uCallbackMessage As Long
     hIcon As Long
     szTip As String * 64
End Type

' tray Return values
Public Const trayLBUTTONDOWN = 7695
Public Const trayLBUTTONUP = 7710
Public Const trayLBUTTONDBLCLK = 7725

Public Const trayRBUTTONDOWN = 7740
Public Const trayRBUTTONUP = 7755
Public Const trayRBUTTONDBLCLK = 7770

Public Const trayMOUSEMOVE = 7680

Public Const WM_RBUTTONUP = &H205
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_LBUTTONDBLCLK = &H203

Global Const NIM_ADD = &H0& 'constants & flags for NotifyIcons
Global Const NIM_MODIFY = &H1
Global Const NIM_DELETE = &H2
Global Const NIF_MESSAGE = &H1
Global Const NIF_ICON = &H2
Global Const NIF_TIP = &H4
Global Const WM_MOUSEMOVE = &H200

Global NI As NOTIFYICONDATA

'Systray API
Declare Function Shell_NotifyIconA Lib "SHELL32" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Public Function ShowProgramInTray()
    With frmMain
        INTRAY = True   'Means App is now in Tray
        NI.cbSize = Len(NI) 'set the length of this structure
        NI.hwnd = .TrayIcon.hwnd 'control to receive messages from
        NI.uID = 0 'uniqueID
        NI.uID = NI.uID + 1
        NI.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP 'operation flags
        NI.uCallbackMessage = WM_MOUSEMOVE 'recieve messages from mouse activities
        NI.hIcon = .TrayIcon.Picture  'the location of the icon to display
    ' Change System Tray Icon's Tool Tip Here bt don't delete chr$(0) [its line carriage here]
        NI.szTip = App.ProductName + Chr$(0)  'LoadResString(Language) + Chr$(0)  'the tool tip to display"
        result = Shell_NotifyIconA(NIM_ADD, NI) 'add the icon to the system tray
    End With
End Function

Public Sub DeleteIcon(pic As Control)
INTRAY = False  'Means app is unloaded or Max mode
    NI.uID = 0 'uniqueID
    NI.uID = NI.uID + 1
    NI.cbSize = Len(NI)
    NI.hwnd = pic.hwnd
    NI.uCallbackMessage = WM_MOUSEMOVE
    result = Shell_NotifyIconA(NIM_DELETE, NI)
End Sub

Public Function NoSysIcon(maxIcon As Boolean)
    On Error Resume Next
    With frmMain
        Select Case maxIcon
            Case False   'Case App in Min Mode
                .Visible = False
                ShowProgramInTray               'Now show TrayIcon PictureBox's Picture in SysTray as icon
            Case Else   'Case App in Max Mode
                .Visible = True
                DeleteIcon .TrayIcon
        End Select
    End With
End Function


