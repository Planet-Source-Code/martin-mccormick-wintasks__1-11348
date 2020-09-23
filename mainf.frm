VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form mainf 
   BorderStyle     =   0  'None
   Caption         =   "Win"
   ClientHeight    =   2340
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2385
   Icon            =   "mainf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   2385
   ShowInTaskbar   =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Timer auto 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   1920
      Top             =   480
   End
   Begin VB.Timer Timer15 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   1920
      Top             =   2760
   End
   Begin VB.Timer Timer5 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1920
      Top             =   2760
   End
   Begin VB.PictureBox pctPrg 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00008000&
      ForeColor       =   &H0000FF00&
      Height          =   270
      Left            =   0
      ScaleHeight     =   18
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   200
      TabIndex        =   6
      Top             =   1080
      Width           =   3000
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   4920
      Top             =   840
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   3840
      Top             =   720
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   5040
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5760
      Top             =   1080
   End
   Begin VB.Label Window 
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   1800
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   "Mouse is over:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1500
      Width           =   2295
   End
   Begin VB.Line Line1 
      Index           =   1
      X1              =   0
      X2              =   3480
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Label Label4 
      Caption         =   "CPU Usage:"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Minutes in Windows:"
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   3375
   End
   Begin VB.Label dates 
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "Date:"
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   360
      Width           =   375
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   -120
      X2              =   3360
      Y1              =   0
      Y2              =   0
   End
   Begin VB.Label Times 
      Height          =   255
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Time:"
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   375
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu about 
         Caption         =   "About..."
         Shortcut        =   ^A
      End
      Begin VB.Menu exer 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu comm 
      Caption         =   "&Controls"
      Begin VB.Menu comp 
         Caption         =   "Windows"
         Begin VB.Menu sd 
            Caption         =   "Shut Down"
            Shortcut        =   {F2}
         End
         Begin VB.Menu re 
            Caption         =   "Restart"
            Shortcut        =   {F3}
         End
         Begin VB.Menu lo 
            Caption         =   "Log Off"
            Shortcut        =   {F4}
         End
         Begin VB.Menu sep99 
            Caption         =   "-"
         End
         Begin VB.Menu Rw 
            Caption         =   "Refresh Windows"
            Shortcut        =   {F11}
         End
         Begin VB.Menu scputer 
            Caption         =   "Secure Computer"
            Shortcut        =   {F8}
         End
         Begin VB.Menu amm 
            Caption         =   "Auto Mouse Move"
            Shortcut        =   ^U
         End
         Begin VB.Menu erb 
            Caption         =   "Empty Recycling Bin"
         End
         Begin VB.Menu set 
            Caption         =   "Set"
            Begin VB.Menu sr 
               Caption         =   "Screen Resoluton"
               Shortcut        =   +^{F1}
            End
            Begin VB.Menu ss123 
               Caption         =   "Screen Saver"
               Begin VB.Menu son 
                  Caption         =   "On"
                  Checked         =   -1  'True
               End
               Begin VB.Menu soff 
                  Caption         =   "Off"
               End
            End
            Begin VB.Menu taskvis 
               Caption         =   "Taskbar"
               Begin VB.Menu vistes 
                  Caption         =   "Show"
                  Checked         =   -1  'True
               End
               Begin VB.Menu hidyis 
                  Caption         =   "Hide"
               End
            End
            Begin VB.Menu di123 
               Caption         =   "Desktop Icons"
               Begin VB.Menu dshow 
                  Caption         =   "Show"
                  Checked         =   -1  'True
               End
               Begin VB.Menu dhide 
                  Caption         =   "Hide"
               End
            End
            Begin VB.Menu acd 
               Caption         =   "Alt - Ctrl - Del"
               Begin VB.Menu ena 
                  Caption         =   "Enabled"
                  Checked         =   -1  'True
               End
               Begin VB.Menu dis 
                  Caption         =   "Disabled"
               End
            End
            Begin VB.Menu ocr 
               Caption         =   "Open CD-Rom"
            End
         End
         Begin VB.Menu sep2000 
            Caption         =   "-"
         End
         Begin VB.Menu ma 
            Caption         =   "Minimize All"
            Shortcut        =   ^M
         End
         Begin VB.Menu wp123 
            Caption         =   "Programs"
            Begin VB.Menu fff 
               Caption         =   "Find Files or Folders..."
               Shortcut        =   ^F
            End
            Begin VB.Menu winfi 
               Caption         =   "WinFiles"
               Shortcut        =   ^W
            End
            Begin VB.Menu re1 
               Caption         =   "Explore..."
               Shortcut        =   ^E
            End
            Begin VB.Menu sep48 
               Caption         =   "-"
            End
            Begin VB.Menu enum 
               Caption         =   "Enumerator"
            End
            Begin VB.Menu sfinder 
               Caption         =   """*"" Finder"
            End
            Begin VB.Menu atd 
               Caption         =   "ASCII to Decimal"
            End
            Begin VB.Menu vc 
               Caption         =   "Veda Creator"
            End
         End
      End
      Begin VB.Menu internet 
         Caption         =   "Internet"
         Begin VB.Menu email 
            Caption         =   "E-mail"
            Begin VB.Menu ce 
               Caption         =   "Check Email..."
               Shortcut        =   ^C
            End
            Begin VB.Menu se123 
               Caption         =   "Send Email..."
               Shortcut        =   ^S
            End
         End
         Begin VB.Menu browser 
            Caption         =   "Internet Browser..."
            Shortcut        =   ^B
         End
         Begin VB.Menu con 
            Caption         =   "Connect"
         End
         Begin VB.Menu dcon 
            Caption         =   "Diconnect"
         End
         Begin VB.Menu ip2 
            Caption         =   "Internet Properties..."
            Shortcut        =   ^P
         End
      End
      Begin VB.Menu cp 
         Caption         =   "Control Panel"
         Begin VB.Menu arp 
            Caption         =   "Add/Remove Programs..."
         End
         Begin VB.Menu std1 
            Caption         =   "Set Time/Date..."
         End
         Begin VB.Menu rs123 
            Caption         =   "Regional Settings..."
         End
         Begin VB.Menu anh 
            Caption         =   "Add new hardware..."
         End
         Begin VB.Menu disp4 
            Caption         =   "Display Properties..."
         End
         Begin VB.Menu ip1 
            Caption         =   "Internet Properties..."
         End
         Begin VB.Menu kp 
            Caption         =   "Keyboard Properties..."
         End
         Begin VB.Menu mp 
            Caption         =   "Mouse Properties..."
         End
         Begin VB.Menu mp2 
            Caption         =   "Modem Properties..."
         End
         Begin VB.Menu sysp 
            Caption         =   "System Properties..."
         End
         Begin VB.Menu np 
            Caption         =   "Network Properties..."
         End
         Begin VB.Menu pp 
            Caption         =   "Password Properties..."
         End
         Begin VB.Menu sp123 
            Caption         =   "Sounds Properties..."
         End
      End
      Begin VB.Menu files 
         Caption         =   "Files"
         Begin VB.Menu copyfile 
            Caption         =   "Copy"
         End
         Begin VB.Menu run 
            Caption         =   "Run..."
            Shortcut        =   ^R
         End
      End
      Begin VB.Menu tp 
         Caption         =   "This Program"
         Begin VB.Menu fot 
            Caption         =   "Force On top"
         End
         Begin VB.Menu aot 
            Caption         =   "Always on top"
            Checked         =   -1  'True
         End
         Begin VB.Menu size 
            Caption         =   "Size"
            Begin VB.Menu compact 
               Caption         =   "Compact"
               Checked         =   -1  'True
            End
            Begin VB.Menu full 
               Caption         =   "Medium"
            End
            Begin VB.Menu full2 
               Caption         =   "Full"
            End
         End
      End
   End
End
Attribute VB_Name = "mainf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4
Private Const SPI_SETSCREENSAVEACTIVE = 17
Private Const SPIF_UPDATEINIFILE = &H1
Private Const SPIF_SENDWININICHANGE = &H2
Const Internet_Autodial_Force_Unattended As Long = 2
Private Declare Function InternetAutodial Lib "wininet.dll" (ByVal dwFlags As Long, ByVal dwReserved As Long) As Long
Private Declare Function InternetAutodialHangup Lib "wininet.dll" (ByVal dwReserved As Long) As Long
Private Declare Function SystemParametersInfo Lib "user32" _
Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal _
uParam As Long, ByVal lpvParam As Long, ByVal fuWinIni As _
Long) As Long
Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Long, ByVal hwndCallback As Long) As Long

Private Declare Function SetCursorPos Lib "user32.dll" (ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private XT(1) As Single, YT As Single, M As Single
Private XScreen As Single, YScreen As Single
Private i As Integer, II As Integer
Private Const MAX_DELOCATION = 250
Private Const PUPIL_DISTANCE = 30
Option Explicit
Private CPU As New CPUUsage
Private Avg As Long                         ' Average of CPU Usage
Private Sum As Long
Private Index As Long
Dim TimeVal
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Const conHwndTopmost = -1
Const conHwndNoTopmost = -2
Const conSwpNoActivate = &H10
Const conSwpShowWindow = &H40
'Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByVal lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Const KEYEVENTF_KEYUP = &H2
Const VK_LWIN = &H5B
Private Const EWX_SHUTDOWN As Long = 1
Private Declare Function ExitWindowsEx Lib "user32" (ByVal dwOptions As Long, ByVal dwReserved As Long) As Long
Private Const EWX_REBOOT As Long = 2
Private Const EWX_LOGOFF As Long = 0
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Const conSwNormal = 1
Private Const SW_SHOW = 5
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Dim dgf
Const SPI_SCREENSAVERRUNNING = 97
Private Declare Function BringWindowToTop Lib "user32.dll" (ByVal hwnd As Long) As Long
Const SWP_HIDEWINDOW = &H80
Const SWP_SHOWWINDOW = &H40
Private Type NOTIFYICONDATA
   cbSize As Long
   hwnd As Long
   uId As Long
   uFlags As Long
   uCallBackMessage As Long
   hIcon As Long
   szTip As String * 64
End Type

'Declare the constants for the API function. These constants can be
'found in the header file Shellapi.h.

'The following constants are the messages sent to the
'Shell_NotifyIcon function to add, modify, or delete an icon from the
'taskbar status area.
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2

'The following constant is the message sent when a mouse event occurs
'within the rectangular boundaries of the icon in the taskbar status
'area.
Private Const WM_MOUSEMOVE = &H200

'The following constants are the flags that indicate the valid
'members of the NOTIFYICONDATA data type.
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4

'The following constants are used to determine the mouse input on the
'the icon in the taskbar status area.

'Left-click constants.
Private Const WM_LBUTTONDBLCLK = &H203   'Double-click
Private Const WM_LBUTTONDOWN = &H201     'Button down
Private Const WM_LBUTTONUP = &H202       'Button up

'Right-click constants.
Private Const WM_RBUTTONDBLCLK = &H206   'Double-click
Private Const WM_RBUTTONDOWN = &H204     'Button down
Private Const WM_RBUTTONUP = &H205       'Button up

'Declare the API function call.
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean

'Dimension a variable as the user-defined data type.
Dim nid As NOTIFYICONDATA







Private Sub about_Click()
Dim dddw
dddw = MsgBox("This program was written by Martin McCormick", vbOKOnly + vbInformation, "About")
End Sub

Private Sub amm_Click()
If amm.Checked = False Then
auto.Enabled = True
amm.Checked = True
Else
auto.Enabled = False
amm.Checked = False
End If
End Sub

'''
Private Sub anh_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl @1", 5)
End Sub

Private Sub aot_Click()
If aot.Checked = True Then
SetWindowPos hwnd, conHwndNoTopmost, 100, 100, 205, 141, conSwpNoActivate Or conSwpShowWindow
aot.Checked = False
Me.Left = 0
Me.Top = 0
mainf.Height = 615
mainf.Width = 2070
GoTo lll:
End If
If aot.Checked = False Then
aot.Checked = True
SetWindowPos hwnd, conHwndTopmost, 0, 0, 205, 141, conSwpNoActivate Or conSwpShowWindow
Me.Left = 0
Me.Top = 0
mainf.Height = 615
mainf.Width = 2070
End If
lll:
End Sub

Private Sub arp_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL appwiz.cpl,,1", 5)
End Sub

Private Sub atd_Click()
Covert.Show
End Sub

Private Sub auto_Timer()
Dim retvals
retvals = SetCursorPos(Rnd * 1000, Rnd * 700)
End Sub

Private Sub browser_Click()
ShellExecute hwnd, "open", "", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub ce_Click()
Shell ("C:\Program Files\Outlook Express\msimn.exe")
End Sub

Private Sub compact_Click()
Timer5.Enabled = False
Timer3.Enabled = False
compact.Checked = True
full2.Checked = False
full.Checked = False
mainf.Height = 615
mainf.Width = 2070
End Sub

Private Sub con_Click()
Dim lResult As Long
lResult = InternetAutodial(Internet_Autodial_Force_Unattended, 0&)
End Sub

Private Sub copyfile_Click()
Form5.Show
End Sub

Private Sub dcon_Click()
Dim lResult As Long
lResult = InternetAutodialHangup(0&)
End Sub

Private Sub dhide_Click()
Dim hwnd As Long
hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hwnd, 0
dhide.Checked = True
dshow.Checked = False
End Sub

Private Sub dis_Click()
callme (True)
ena.Checked = False
dis.Checked = True
End Sub

Private Sub disp4_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL desk.cpl,,0", 5)
End Sub

Private Sub dshow_Click()
Dim hwnd As Long
hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hwnd, 5
dhide.Checked = False
dshow.Checked = True
End Sub

Private Sub ena_Click()
callme (False)
ena.Checked = True
dis.Checked = False
End Sub

Private Sub enum_Click()
Form1.Show
End Sub

Private Sub erb_Click()
Dim retvaL
retvaL = SHEmptyRecycleBin(Form1.hwnd, "", SHERB_NOPROGRESSUI)

End Sub

Private Sub exer_Click()
Unload Me
End Sub

Private Sub fff_Click()
Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(70, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub Form_Load()

mainf.Height = 615
mainf.Width = 2070
 XScreen = Screen.Width / Screen.TwipsPerPixelX
    YScreen = Screen.Height / Screen.TwipsPerPixelY
    II = 1
SetWindowPos hwnd, conHwndTopmost, 0, 0, 205, 141, conSwpNoActivate Or conSwpShowWindow
'CPU.InitCPUUsage
Me.Left = 0
Me.Top = 0
mainf.Height = 615
mainf.Width = 2070
nid.cbSize = Len(nid)
   nid.hwnd = mainf.hwnd
   nid.uId = vbNull
   nid.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
   nid.uCallBackMessage = WM_MOUSEMOVE
   nid.hIcon = mainf.Icon
   nid.szTip = "Windows Control Program" & vbNullChar

   'Call the Shell_NotifyIcon function to add the icon to the taskbar
   'status area.
   Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    'Event occurs when the mouse pointer is within the rectangular
    'boundaries of the icon in the taskbar status area.
    Dim Msg As Long
    Dim sFilter As String
    Msg = X / Screen.TwipsPerPixelX
    Select Case Msg
       Case WM_LBUTTONDOWN
       Case WM_LBUTTONUP
       mainf.Visible = True
       mainf.WindowState = 0
       AppActivate ("Win")
       Case WM_LBUTTONDBLCLK
      
       mainf.Visible = True
       mainf.WindowState = 0
       AppActivate ("Win")
       Case WM_RBUTTONDOWN
          Dim ToolTipString As String
           
          If ToolTipString <> "" Then
             nid.szTip = ToolTipString & vbNullChar
             Shell_NotifyIcon NIM_MODIFY, nid
          End If
       Case WM_RBUTTONUP
       Case WM_RBUTTONDBLCLK
    End Select

End Sub



Private Sub Form_Unload(Cancel As Integer)
If dgf = 1 Then
Cancel = 1
Else
Dim qwe
qwe = MsgBox("Are you sure you want to close Win Controls?", vbQuestion + vbYesNo, "Are you sure?")
If qwe = vbYes Then GoTo h:
If qwe = vbNo Then
Cancel = 1
GoTo h2:
End If
End If
h:
Shell_NotifyIcon NIM_DELETE, nid
h2:
End Sub

Private Sub fot_Click()
If fot.Checked = False Then
Timer15.Enabled = True
fot.Checked = True
GoTo lll:
Else
fot.Checked = False
Timer15.Enabled = False
End If


lll:
End Sub

Private Sub full_Click()
full2.Checked = False
full.Checked = True
compact.Checked = False
mainf.Height = 2000
mainf.Width = 2500
Timer3.Enabled = True
Timer5.Enabled = False
End Sub

Private Sub full2_Click()
Timer3.Enabled = True
mainf.Height = 2900
mainf.Width = 2500
full2.Checked = True
full.Checked = False
compact.Checked = False
Timer5.Enabled = True
End Sub

Private Sub hidyis_Click()
Dim rtn As Long
rtn = FindWindow("Shell_traywnd", "")
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
vistes.Checked = False
hidyis.Checked = True
End Sub

Private Sub ip1_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5)
End Sub

Private Sub ip2_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl,,0", 5)
End Sub

Private Sub kp_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @1", 5)
End Sub

Private Sub lo_Click()
Dim abc
abc = MsgBox("Are you sure you want to log off the computer?", vbYesNo + vbQuestion, "Log Off")
If abc = vbYes Then
Dim lngresult
lngresult = ExitWindowsEx(EWX_LOGOFF, 0&)
End If

End Sub

Private Sub ma_Click()
Call keybd_event(VK_LWIN, 0, 0, 0)
Call keybd_event(77, 0, 0, 0)
Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub mp_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL main.cpl @0", 5)
End Sub

Private Sub mp2_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL modem.cpl", 5)
End Sub

Private Sub np_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL netcpl.cpl", 5)
End Sub

Private Sub ocr_Click()
Dim lngReturn As Long
Dim strReturn As Long
lngReturn = mciSendString("set CDAudio door open", strReturn, 127, 0)
End Sub

Private Sub pp_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL password.cpl", 5)
End Sub

Private Sub re_Click()
Dim abc
abc = MsgBox("Are you sure you want to restart the computer?", vbYesNo + vbQuestion, "Restart")
If abc = vbYes Then
Dim lngresult
lngresult = ExitWindowsEx(EWX_REBOOT, 0&)
End If
End Sub

Private Sub re1_Click()
Call keybd_event(VK_LWIN, 0, 0, 0)
    Call keybd_event(69, 0, 0, 0)
    Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
End Sub

Private Sub rs123_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL intl.cpl,,0", 5)
End Sub

Private Sub run_Click()
runf.Show
End Sub

Private Sub Rw_Click()
Dim rtn As Long
rtn = FindWindow("Shell_traywnd", "")
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
callme (False)
Dim hwnd As Long
hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hwnd, 5
comm.Enabled = True
file.Enabled = True
End Sub

Private Sub scputer_Click()
Dim hwnd As Long
hwnd = FindWindowEx(0&, 0&, "Progman", vbNullString)
ShowWindow hwnd, 0
Dim rtn As Long
rtn = FindWindow("Shell_traywnd", "")
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_HIDEWINDOW)
Call keybd_event(VK_LWIN, 0, 0, 0)
Call keybd_event(77, 0, 0, 0)
Call keybd_event(VK_LWIN, 0, KEYEVENTF_KEYUP, 0)
comm.Enabled = False
file.Enabled = False
dgf = 1
Timer1.Enabled = True
callme (True)
End Sub

Private Sub sd_Click()
Dim abc
abc = MsgBox("Are you sure you want to shut down the computer?", vbYesNo + vbQuestion, "Shutdown")
If abc = vbYes Then
Dim lngresult
lngresult = ExitWindowsEx(EWX_SHUTDOWN, 0&)
End If
End Sub

Private Sub se123_Click()
ShellExecute hwnd, "open", "mailto:", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub sfinder_Click()
frmPassword.Show
End Sub

Private Sub soff_Click()
ToggleScreenSaverActive (False)
son.Checked = False
soff.Checked = True
End Sub

Private Sub son_Click()
ToggleScreenSaverActive (True)
son.Checked = True
soff.Checked = False
End Sub

Private Sub sp123_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL mmsys.cpl @1", 5)
End Sub

Private Sub sr_Click()
fresolution.Show
End Sub

Private Sub std1_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL timedate.cpl", 5)
End Sub

Private Sub sysp_Click()
Dim dblreturn
dblreturn = Shell("rundll32.exe shell32.dll,Control_RunDLL sysdm.cpl,,0", 5)
End Sub
Private Sub callme(huh As Boolean)
Dim gd
gd = SystemParametersInfo(97, huh, CStr(1), 0)
End Sub


Private Sub Timer1_Timer()
Dim fgt
fgt = InputBox("Please enter access password:", "Enter Password")
If fgt = "vbzm105" Then
dgf = 0
Timer1.Enabled = False
Call enabler
End If
End Sub

Private Sub Timer15_Timer()
SetWindowPos hwnd, conHwndTopmost, 0, 0, 138, 41, conSwpNoActivate Or conSwpShowWindow

End Sub

Private Sub Timer2_Timer()
If mainf.WindowState = 1 Then mainf.Visible = False
End Sub

Private Sub Timer3_Timer()
   
    
     Dim tmp As Long
    tmp = CPU.GetCPUUsage
    Sum = Sum + tmp
    Index = Index + 1
    Avg = Int(Sum / Index)
    'Draw the bar
    pctPrg.Cls
    pctPrg.Line (0, 0)-(tmp, 18), , BF
    pctPrg.Line (Avg, 0)-(Avg, 18), &HFF
    pctPrg.Line (Avg + 1, 0)-(Avg + 1, 18), &HFF
    DoEvents
dates.Caption = Format(Date, "mm/dd/yyyy")
Times.Caption = Format(Time, "hh:mm:ss")
Dim lngTickCount As Long
lngTickCount = GetTickCount
Label3.Caption = CStr(Round((lngTickCount / 1000 / 60))) & " Minutes in Windows"
End Sub

Private Sub Timer5_Timer()
Dim cp As POINTAPI, hwnd As Long, s As String
    GetCursorPos cp
     hwnd = WindowFromPoint(cp.X, cp.Y)
    s = Space(128)
    GetWindowText hwnd, s, 128
    If Asc(Left(s, 1)) = 0 Then GetClassName hwnd, s, 128
    Window.Caption = s
    DoEvents
End Sub

Private Sub vc_Click()
veda.Show
End Sub

Private Sub vistes_Click()
Dim rtn As Long
rtn = FindWindow("Shell_traywnd", "")
Call SetWindowPos(rtn, 0, 0, 0, 0, 0, SWP_SHOWWINDOW)
vistes.Checked = True
hidyis.Checked = False
End Sub
Private Sub enabler()
Call Rw_Click
End Sub

Private Sub winfi_Click()
Shell ("C:\windows\winfile.exe")
End Sub
Public Function ToggleScreenSaverActive(Active As Boolean) _
   As Boolean
Dim lActiveFlag As Long
Dim retvaL As Long

lActiveFlag = IIf(Active, 1, 0)
retvaL = SystemParametersInfo(SPI_SETSCREENSAVEACTIVE, _
   lActiveFlag, 0, 0)
ToggleScreenSaverActive = retvaL > 0

End Function


