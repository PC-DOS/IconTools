VERSION 5.00
Begin VB.Form frmConvertMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pic To Icon - PC-DOS Workshop"
   ClientHeight    =   7815
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   14340
   Icon            =   "frmConvertMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7815
   ScaleWidth      =   14340
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picTransparentColour 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   420
      Left            =   8310
      ScaleHeight     =   390
      ScaleWidth      =   3945
      TabIndex        =   20
      Top             =   4455
      Width           =   3975
   End
   Begin VB.PictureBox picsource 
      Height          =   4095
      Left            =   5460
      ScaleHeight     =   4035
      ScaleWidth      =   5265
      TabIndex        =   19
      Top             =   7635
      Visible         =   0   'False
      Width           =   5325
   End
   Begin VB.CommandButton Command2 
      Caption         =   "初始化(&I)"
      Height          =   1260
      Left            =   7560
      Picture         =   "frmConvertMain.frx":0ECA
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6435
      Width           =   6675
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存为ICO文件(&S)"
      Height          =   1260
      Left            =   7560
      Picture         =   "frmConvertMain.frx":1D94
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5115
      Width           =   6675
   End
   Begin VB.Frame Frame2 
      Caption         =   "转换"
      Height          =   4050
      Left            =   7455
      TabIndex        =   7
      Top             =   975
      Width           =   6810
      Begin VB.CommandButton cmdPickColour 
         BackColor       =   &H80000005&
         Caption         =   "選擇..."
         Height          =   390
         Left            =   4920
         TabIndex        =   23
         Top             =   3495
         Width           =   1755
      End
      Begin VB.Frame Frame3 
         Caption         =   "大小"
         Enabled         =   0   'False
         Height          =   3180
         Left            =   4095
         TabIndex        =   10
         Top             =   165
         Width           =   2595
         Begin VB.OptionButton Option6 
            Caption         =   "&24*24"
            Enabled         =   0   'False
            Height          =   405
            Left            =   150
            TabIndex        =   21
            Top             =   801
            Width           =   2190
         End
         Begin VB.OptionButton Option1 
            Caption         =   "&16*16"
            Enabled         =   0   'False
            Height          =   405
            Left            =   150
            TabIndex        =   15
            Top             =   375
            Width           =   2190
         End
         Begin VB.OptionButton Option2 
            Caption         =   "&32*32"
            Enabled         =   0   'False
            Height          =   405
            Left            =   150
            TabIndex        =   14
            Top             =   1227
            Value           =   -1  'True
            Width           =   2190
         End
         Begin VB.OptionButton Option3 
            Caption         =   "&48*48"
            Enabled         =   0   'False
            Height          =   405
            Left            =   150
            TabIndex        =   13
            Top             =   1653
            Width           =   2190
         End
         Begin VB.OptionButton Option4 
            Caption         =   "&64*64"
            Enabled         =   0   'False
            Height          =   405
            Left            =   150
            TabIndex        =   12
            Top             =   2079
            Width           =   2190
         End
         Begin VB.OptionButton Option5 
            Caption         =   "&128*128"
            Enabled         =   0   'False
            Height          =   405
            Left            =   150
            TabIndex        =   11
            Top             =   2505
            Width           =   2190
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   3150
         Left            =   75
         ScaleHeight     =   206
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   259
         TabIndex        =   8
         Top             =   180
         Width           =   3945
         Begin VB.PictureBox Picture2 
            Height          =   300
            Left            =   1950
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   9
            Top             =   1200
            Width           =   300
         End
      End
      Begin VB.Label lblTransparentColour 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "透明色"
         Height          =   180
         Left            =   150
         TabIndex        =   22
         Top             =   3600
         Width           =   540
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "轉換圖像源"
      Height          =   6735
      Left            =   105
      TabIndex        =   1
      Top             =   975
      Width           =   7305
      Begin VB.PictureBox PicturePreview 
         Enabled         =   0   'False
         Height          =   5730
         Left            =   120
         ScaleHeight     =   5670
         ScaleWidth      =   6750
         TabIndex        =   6
         Top             =   645
         Width           =   6810
         Begin VB.Image Image1 
            Enabled         =   0   'False
            Height          =   2865
            Left            =   60
            MouseIcon       =   "frmConvertMain.frx":2C5E
            MousePointer    =   2  'Cross
            Top             =   60
            Width           =   4695
         End
      End
      Begin VB.VScrollBar VScroll1 
         Enabled         =   0   'False
         Height          =   5715
         Left            =   6945
         TabIndex        =   5
         Top             =   630
         Width           =   255
      End
      Begin VB.HScrollBar HScroll1 
         Enabled         =   0   'False
         Height          =   255
         Left            =   105
         TabIndex        =   4
         Top             =   6360
         Width           =   6825
      End
      Begin VB.CommandButton CommandBrowse 
         Caption         =   "瀏覽(&B)..."
         Height          =   360
         Left            =   6090
         TabIndex        =   3
         Top             =   240
         Width           =   1110
      End
      Begin VB.TextBox txtUrl 
         Height          =   345
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   240
         Width           =   5925
      End
   End
   Begin VB.Image IC 
      Height          =   495
      Left            =   6570
      Top             =   3660
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "透過[瀏覽]按鈕載入圖像，在[大小]中指定圖標大小，之後保存即可"
      Height          =   390
      Index           =   1
      Left            =   960
      TabIndex        =   17
      Top             =   555
      Width           =   13290
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmConvertMain.frx":30A0
      Height          =   390
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   195
      Width           =   13290
   End
   Begin VB.Image ImageIcon 
      Height          =   720
      Left            =   105
      Picture         =   "frmConvertMain.frx":316C
      Top             =   120
      Width           =   720
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuOpen 
         Caption         =   "打開(&O)..."
         Shortcut        =   ^O
      End
      Begin VB.Menu b1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuInit 
         Caption         =   "初始化(&I)"
         Shortcut        =   ^I
      End
      Begin VB.Menu b2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出"
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "工具(&T)"
      Begin VB.Menu mnuGet 
         Caption         =   "從Win32 PE文件提取圖標資源(&G)..."
         Shortcut        =   ^G
      End
      Begin VB.Menu b3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuWizard 
         Caption         =   "圖標創建向導(&W)..."
         Shortcut        =   ^W
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "幫助(&H)"
      Begin VB.Menu mnuAbout 
         Caption         =   "關於Pic To Icon(&A)..."
      End
   End
End
Attribute VB_Name = "frmConvertMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const DEFAULT_SIZE_VALUE = 810
Dim CommonDialog1 As New CCommonDialog
 Private Type PROCESSENTRY32
 dwSize As Long
 cntUsage As Long
 th32ProcessID As Long '进程ID
 th32DefaultHeapID As Long '堆栈ID
 th32ModuleID As Long '模块ID
 cntThreads As Long
 th32ParentProcessID As Long '父进程ID
 pcPriClassBase As Long
 dwFlags As Long
 szExeFile As String * 260
 End Type
 Private Type MEMORYSTATUS
 dwLength As Long
 dwMemoryLoad As Long
 dwTotalPhys As Long
 dwAvailPhys As Long
 dwTotalPageFile As Long
 dwAvailPageFile As Long
 dwTotalVirtual As Long
 dwAvailVirtual As Long
 End Type
 Private Declare Function NtQuerySystemInformation Lib "ntdll" (ByVal dwInfoType As Long, ByVal lpStructure As Long, ByVal dwSize As Long, ByVal dwReserved As Long) As Long
Private Const SYSTEM_BASICINFORMATION = 0&
Private Const SYSTEM_PERFORMANCEINFORMATION = 2&
Private Const SYSTEM_TIMEINFORMATION = 3&
Private Const NO_ERROR = 0
Private Type LARGE_INTEGER
    dwLow As Long
    dwHigh As Long
End Type

Private Type SYSTEM_PERFORMANCE_INFORMATION
    liIdleTime As LARGE_INTEGER
    dwSpare(0 To 75) As Long
End Type
Private Type SYSTEM_BASIC_INFORMATION
    dwUnknown1 As Long
    uKeMaximumIncrement As Long
    uPageSize As Long
    uMmNumberOfPhysicalPages As Long
    uMmLowestPhysicalPage As Long
    uMmHighestPhysicalPage As Long
    uAllocationGranularity As Long
    pLowestUserAddress As Long
    pMmHighestUserAddress As Long
    uKeActiveProcessors As Long
    bKeNumberProcessors As Byte
    bUnknown2 As Byte
    wUnknown3 As Integer
End Type
Private Type SYSTEM_TIME_INFORMATION
    liKeBootTime As LARGE_INTEGER
    liKeSystemTime As LARGE_INTEGER
    liExpTimeZoneBias As LARGE_INTEGER
    uCurrentTimeZoneId As Long
    dwReserved As Long
End Type

Private lidOldIdle As LARGE_INTEGER
Private liOldSystem As LARGE_INTEGER
 Private Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
 Private Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long '获取首个进程
 Private Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long '获取下个进程
 Private Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, lppe As PROCESSENTRY32) As Long '释放句柄
 Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)
 Private Const TH32CS_SNAPPROCESS = &H2&
 Private Const KEYEVENTF_EXTENDEDKEY = &H1
Private Const KEYEVENTF_KEYUP = &H2
Dim IsHideToTray As Boolean
Private Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100
Private Const FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
Private Const FORMAT_MESSAGE_FROM_HMODULE = &H800
Private Const FORMAT_MESSAGE_FROM_STRING = &H400
Private Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Private Const FORMAT_MESSAGE_IGNORE_INSERTS = &H200
Private Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Const VK_LWIN = &H5B
Private Const WM_KEYUP = &H101
Private Const WM_KEYDOWN = &H100
Private Declare Function DebugActiveProcess Lib "kernel32" (ByVal dwProcessId As Long) As Long
Private Declare Sub DebugBreak Lib "kernel32" ()
Private Const SM_DEBUG = 22
Private Const DEBUG_ONLY_THIS_PROCESS = &H2
Private Const DEBUG_PROCESS = &H1
Private Type USER_DIALOG_CONFIG
lpTitle As String
lpIcon As Integer
lpMessage As String
End Type
Private Type USER_APP_RUN
lpAppPath As String
lpAppParam As String
lpRunMode As Integer
End Type
Private Type APP_TASK_PARAM
lpTimerType As Integer
lpDelay As Long
lpRunHour As Integer
lpRunMinute As Integer
lpRunSecond As Integer
lpCurrentHour As Integer
lpCurrentMinute As Integer
lpCurrentSecond As Integer
lpTaskEnum As Integer
lpTaskFriendlyDisplayName As String
lpRunning As Boolean
End Type
Dim lpDialogCfg As USER_DIALOG_CONFIG
Dim lpAppCfg As USER_APP_RUN
Dim lpTaskCfg As APP_TASK_PARAM
Const SC_SCREENSAVE = &HF140&
Dim IsCodeUse As Boolean
Private Declare Function WinExec Lib "kernel32" (ByVal lpCmdLine As String, ByVal nCmdShow As Long) As Long
Private Declare Function AppendMenu Lib "user32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const GWL_WNDPROC = (-4)
Private Const WM_SYSCOMMAND = &H112
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SetFileAttributes Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpFileName As String, ByVal dwFileAttributes As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Const FILE_ATTRIBUTE_ARCHIVE = &H20
Private Const FILE_ATTRIBUTE_COMPRESSED = &H800
Private Const FILE_ATTRIBUTE_DIRECTORY = &H10
Private Const FILE_ATTRIBUTE_HIDDEN = &H2
Private Const FILE_ATTRIBUTE_NORMAL = &H80
Private Const FILE_ATTRIBUTE_READONLY = &H1
Private Const FILE_ATTRIBUTE_SYSTEM = &H4
Private Const FILE_ATTRIBUTE_TEMPORARY = &H100
Dim lpSize As Long
Dim bchk As Boolean
Dim lpFilePath As String
Const READ_CONTROL = &H20000
Const KEY_QUERY_VALUE = &H1
Const KEY_SET_VALUE = &H2
Const KEY_CREATE_SUB_KEY = &H4
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const KEY_NOTIFY = &H10
Const KEY_CREATE_LINK = &H20
Const KEY_ALL_ACCESS = KEY_QUERY_VALUE + KEY_SET_VALUE + _
KEY_CREATE_SUB_KEY + KEY_ENUMERATE_SUB_KEYS + _
KEY_NOTIFY + KEY_CREATE_LINK + READ_CONTROL
Const HKEY_LOCAL_MACHINE = &H80000002
Const ERROR_SUCCESS = 0
Const REG_SZ = 1
Const REG_DWORD = 4
Const gREGKEYSYSINFOLOC = "SOFTWARE\Microsoft\Shared Tools Location"
Const gREGVALSYSINFOLOC = "MSINFO"
Const gREGKEYSYSINFO = "SOFTWARE\Microsoft\Shared Tools\MSINFO"
Const gREGVALSYSINFO = "PATH"
Private Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Long) As Long
Const MAX_FILE_SIZE = 1.5 * (1024 ^ 3)
Private Const HWND_BOTTOM = 1
Private Const HWND_BROADCAST = &HFFFF&
Private Const HWND_DESKTOP = 0
Private Const HWND_NOTOPMOST = -2
Private Const WS_EX_TRANSPARENT = &H20&
Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function OffsetRect Lib "user32" (lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Private Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
'很多朋友都见到过能在托盘图标上出现气球提示的软件，不说软件，就是在“磁盘空间不足”时Windows给出的提示就属于气球提示，那么怎样在自己的程序中添加这样的气球提示呢？
   
'其实并不难，关键就在添加托盘图标时所使用的NOTIFYICONDATA结构，源代码如下：
Private Declare Function Shell_NotifyIcon Lib "shell32.dll" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long
   
Private Type NOTIFYICONDATA
cbSize   As Long     '   结构大小(字节)
hwnd   As Long     '   处理消息的窗口的句柄
uID   As Long     '   唯一的标识符
uFlags   As Long     '   Flags
uCallbackMessage   As Long     '   处理消息的窗口接收的消息
hIcon   As Long     '   托盘图标句柄
szTip   As String * 128         '   Tooltip   提示文本
dwState   As Long     '   托盘图标状态
dwStateMask   As Long     '   状态掩码
szInfo   As String * 256         '   气球提示文本
uTimeoutOrVersion   As Long     '   气球提示消失时间或版本
'   uTimeout   -   气球提示消失时间(单位:ms,   10000   --   30000)
'   uVersion   -   版本(0   for   V4,   3   for   V5)
szInfoTitle   As String * 64         '   气球提示标题
dwInfoFlags   As Long     '   气球提示图标
End Type
   
'   dwState   to   NOTIFYICONDATA   structure
Private Const NIS_HIDDEN = &H1           '   隐藏图标
Private Const NIS_SHAREDICON = &H2           '   共享图标
   
'   dwInfoFlags   to   NOTIFIICONDATA   structure
Private Const NIIF_NONE = &H0           '   无图标
Private Const NIIF_INFO = &H1           '   "消息"图标
Private Const NIIF_WARNING = &H2           '   "警告"图标
Private Const NIIF_ERROR = &H3           '   "错误"图标
   
'   uFlags   to   NOTIFYICONDATA   structure
Private Const NIF_ICON       As Long = &H2
Private Const NIF_INFO       As Long = &H10
Private Const NIF_MESSAGE       As Long = &H1
Private Const NIF_STATE       As Long = &H8
Private Const NIF_TIP       As Long = &H4
   
'   dwMessage   to   Shell_NotifyIcon
Private Const NIM_ADD       As Long = &H0
Private Const NIM_DELETE       As Long = &H2
Private Const NIM_MODIFY       As Long = &H1
Private Const NIM_SETFOCUS       As Long = &H3
Private Const NIM_SETVERSION       As Long = &H4
Private Type RECTL
        Left As Long
        TOp As Long
        Right As Long
        Bottom As Long
End Type
Private Type RECT
        Left As Long
        TOp As Long
        Right As Long
        Bottom As Long
End Type
Dim cRect As RECT
Const LCR_UNLOCK = 0
Dim dwMouseFlag As Integer
Const ME_LBCLICK = 1
Const ME_LBDBLCLICK = 2
Const ME_RBCLICK = 3
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_ABSOLUTE = &H8000
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
Private Const MOUSEEVENTF_MIDDLEUP = &H40
Private Const MOUSEEVENTF_MOVE = &H1
Private Const MOUSEEVENTF_RIGHTDOWN = &H8
Private Const MOUSEEVENTF_RIGHTUP = &H10
Private Const MOUSETRAILS = 39
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Const SWP_NOACTIVATE = &H10
Private Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Dim lpszCaptionNew As String
Private Const SC_MINIMIZE = &HF020&
Private Const WS_MAXIMIZEBOX = &H10000
Dim HKStateCtrl As Integer
Dim HKStateFn As Integer
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MINIMIZE = &H20000000
Private Const WS_ICONIC = WS_MINIMIZE
Const SC_ICON = SC_MINIMIZE
Const SC_TASKLIST = &HF130&
Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Dim bCodeUse As Boolean
Private Const WS_CAPTION = &HC00000
Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Const PROCESS_ALL_ACCESS = &H1F0FFF
Const SC_RESTORE = &HF120&
Const TH32CS_SNAPHEAPLIST = &H1
Const TH32CS_SNAPTHREAD = &H4
Const TH32CS_SNAPMODULE = &H8
Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Const TH32CS_INHERIT = &H80000000
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Dim lMeWinStyle As Long
Const SWP_SHOWWINDOW = &H40
Const SWP_HIDEWINDOW = &H80
Const SWP_NOOWNERZORDER = &H200
Const SWP_NOREPOSITION = SWP_NOOWNERZORDER
Private Const SC_MOVE = &HF010&
Private Const SC_SIZE = &HF000&
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Const WS_EX_APPWINDOW = &H40000
Private Type WINDOWINFORMATION
hWindow As Long
hWindowDC As Long
hThreadProcess As Long
hThreadProcessID As Long
lpszCaption As String
lpszClassName As String
lpszThreadProcessName As String * 1024
lpszThreadProcessPath As String
lpszExe As String
lpszPath As String
End Type
Private Type WINDOWPARAM
bEnabled As Boolean
bHide As Boolean
bTrans As Boolean
bClosable As Boolean
bSizable As Boolean
bMinisizable As Boolean
bTop As Boolean
lpTransValue As Integer
End Type
Dim lpWindow As WINDOWINFORMATION
Dim lpWindowParam() As WINDOWPARAM
Dim lpCur As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const WM_GETTEXT = &HD
Private Const WM_GETTEXTLENGTH = &HE
Dim lpRtn As Long
Dim hWindow As Long
Dim lpLength As Long
Dim lpArray() As Byte
Dim lpArray2() As Byte
Dim lpBuff As String
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const LWA_COLORKEY = &H1
Private Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function DeleteMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Const MF_BYPOSITION = &H400&
Private Const MF_REMOVE = &H1000&
Private Const WS_SYSMENU = &H80000
Private Const GWL_STYLE = (-16)
Private Const MF_BYCOMMAND = &H0
Private Const SC_CLOSE = &HF060&
Private Declare Function SetMenu Lib "user32" (ByVal hwnd As Long, ByVal hMenu As Long) As Long
Private Const MF_INSERT = &H0&
Private Const SC_MAXIMIZE = &HF030&
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Type WINDOWINFOBOXDATA
lpszCaption As String
lpszClass As String
lpszThread As String
lpszHandle As String
lpszDC As String
End Type
Dim dwWinInfo As WINDOWINFOBOXDATA
Dim bError As Boolean
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Const WM_CLOSE = &H10
Private Const HWND_TOPMOST = -1
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOREDRAW = &H8
Private Const SWP_NOMOVE = &H2
Dim mov As Boolean
Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Const ANYSIZE_ARRAY = 1
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Private Type LUID
UsedPart As Long
IgnoredForNowHigh32BitPart As Long
End Type
Private Type TOKEN_PRIVILEGES
PrivilegeCount As Long
TheLuid As LUID
Attributes As Long
End Type
Private Declare Function OpenProcessToken Lib "advapi32" (ByVal _
ProcessHandle As Long, _
ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32" _
Alias "LookupPrivilegeValueA" _
(ByVal lpSystemName As String, ByVal lpName As String, lpLuid _
As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32" _
(ByVal TokenHandle As Long, _
ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES _
, ByVal BufferLength As Long, _
PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Type TestCounter
TimesLeft As Integer
ResetTime As Integer
End Type
Dim PassTest As TestCounter
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Type POINTAPI
x As Long
y As Long
End Type
Private Const VK_ADD = &H6B
Private Const VK_ATTN = &HF6
Private Const VK_BACK = &H8
Private Const VK_CANCEL = &H3
Private Const VK_CAPITAL = &H14
Private Const VK_CLEAR = &HC
Private Const VK_CONTROL = &H11
Private Const VK_CRSEL = &HF7
Private Const VK_DECIMAL = &H6E
Private Const VK_DELETE = &H2E
Private Const VK_DIVIDE = &H6F
Private Const VK_DOWN = &H28
Private Const VK_END = &H23
Private Const VK_EREOF = &HF9
Private Const VK_ESCAPE = &H1B
Private Const VK_EXECUTE = &H2B
Private Const VK_EXSEL = &HF8
Private Const VK_F1 = &H70
Private Const VK_F10 = &H79
Private Const VK_F11 = &H7A
Private Const VK_F12 = &H7B
Private Const VK_F13 = &H7C
Private Const VK_F14 = &H7D
Private Const VK_F15 = &H7E
Private Const VK_F16 = &H7F
Private Const VK_F17 = &H80
Private Const VK_F18 = &H81
Private Const VK_F19 = &H82
Private Const VK_F2 = &H71
Private Const VK_F20 = &H83
Private Const VK_F21 = &H84
Private Const VK_F22 = &H85
Private Const VK_F23 = &H86
Private Const VK_F24 = &H87
Private Const VK_F3 = &H72
Private Const VK_F4 = &H73
Private Const VK_F5 = &H74
Private Const VK_F6 = &H75
Private Const VK_F7 = &H76
Private Const VK_F8 = &H77
Private Const VK_F9 = &H78
Private Const VK_HELP = &H2F
Private Const VK_HOME = &H24
Private Const VK_INSERT = &H2D
Private Const VK_LBUTTON = &H1
Private Const VK_LCONTROL = &HA2
Private Const VK_LEFT = &H25
Private Const VK_LMENU = &HA4
Private Const VK_LSHIFT = &HA0
Private Const VK_MBUTTON = &H4
Private Const VK_MENU = &H12
Private Const VK_MULTIPLY = &H6A
Private Const VK_NEXT = &H22
Private Const VK_NONAME = &HFC
Private Const VK_NUMLOCK = &H90
Private Const VK_NUMPAD0 = &H60
Private Const VK_NUMPAD1 = &H61
Private Const VK_NUMPAD2 = &H62
Private Const VK_NUMPAD3 = &H63
Private Const VK_NUMPAD4 = &H64
Private Const VK_NUMPAD5 = &H65
Private Const VK_NUMPAD6 = &H66
Private Const VK_NUMPAD7 = &H67
Private Const VK_NUMPAD8 = &H68
Private Const VK_NUMPAD9 = &H69
Private Const VK_OEM_CLEAR = &HFE
Private Const VK_PA1 = &HFD
Private Const VK_PAUSE = &H13
Private Const VK_PLAY = &HFA
Private Const VK_PRINT = &H2A
Private Const VK_PRIOR = &H21
Private Const VK_PROCESSKEY = &HE5
Private Const VK_RBUTTON = &H2
Private Const VK_RCONTROL = &HA3
Private Const VK_RETURN = &HD
Private Const VK_RIGHT = &H27
Private Const VK_RMENU = &HA5
Private Const VK_RSHIFT = &HA1
Private Const VK_SCROLL = &H91
Private Const VK_SELECT = &H29
Private Const VK_SEPARATOR = &H6C
Private Const VK_SHIFT = &H10
Private Const VK_SNAPSHOT = &H2C
Private Const VK_SPACE = &H20
Private Const VK_SUBTRACT = &H6D
Private Const VK_TAB = &H9
Private Const VK_UP = &H26
Private Const VK_ZOOM = &HFB
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Dim lpX As Long
Dim lpY As Long
Private Type FILEINFO
lpPath As String
lpDateLastChanged As Date
lpAttribList As Integer
lpSize As Long
lpHeader As String * 25
lpType As String
lpAttrib As String
End Type
Dim lpFile As FILEINFO
Public act As Boolean
Dim regsvrvrt
Dim unregsvrvrt
Dim regflag As Boolean
Dim unregflag  As Boolean
Dim ream
Private Type BROWSEINFO
hOwner As Long
pidlRoot As Long
pszDisplayName As String
lpszTitle As String
ulFlags As Long
lpfn As Long
lParam As Long
iImage As Long
End Type
Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_NEWDIALOGSTYLE = &H40
Private Const BIF_NONEWFOLDERBUTTON = &H200
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" _
(ByVal pidl As Long, _
ByVal pszPath As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" _
(lpBrowseInfo As BROWSEINFO) As Long
Private Declare Function CloseScreenFun Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Const SC_MONITORPOWER = &HF170&
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Declare Function OleTranslateColor Lib "OLEPRO32.DLL" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Const CLR_INVALID = -1
Private Declare Function GetPixelAPI Lib "gdi32" Alias "GetPixel" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long

Private m_cSource As cAlphaDIBSection
Private m_lTransparentColour As OLE_COLOR
Private m_pic As StdPicture
Private m_iTab As Long

Private Sub createIconAtSize( _
      cFI As cFileIcon, _
      ByVal lIndex As Long, _
      ByVal lWidth As Long, _
      ByVal lHeight As Long _
   )
Dim cResampled As cAlphaDIBSection
   ' Resample the input bitmap:
   Set cResampled = m_cSource.AlphaResample(lWidth)
   If (cResampled.Height < lHeight) Then
      ' Need to place the item in a new dib of the
      ' correct size:
      Dim cSized As New cAlphaDIBSection
      cSized.Create lWidth, lHeight
      cSized.SetBackgroundColor m_lTransparentColour
      cSized.SetColourTransparent m_lTransparentColour
      cResampled.CopyTo cSized, (lWidth - cResampled.Width) \ 2, (lHeight - cResampled.Height) \ 2
      Set cResampled = cSized
   End If
   
   ' Set the alpha bits to the result
   cFI.SetImageBits lIndex, cResampled.DIBSectionBitsPtr
   
Dim b() As Byte
Dim lWidthBytes As Long
   lWidthBytes = ((cResampled.Width + 31) \ 32) * 4
   ReDim b(0 To lWidthBytes - 1, 0 To lHeight - 1) As Byte
   
   createMask cResampled, b()
   cFI.SetMaskBits lIndex, VarPtr(b(0, 0))
      
End Sub

Private Sub createMask( _
      cDib As cAlphaDIBSection, _
      b() As Byte _
   )
Dim lWidthBytes As Long
Dim lHeight As Long
Dim lCurVal As Long
Dim lBit As Long
Dim x As Long
Dim y As Long
Dim tSA As SAFEARRAY2D
Dim bDib() As Byte
Dim xOut As Long
Dim yOut As Long
         
   ' Get the bits in the from DIB section:
   With tSA
      .cbElements = 1
      .cDims = 2
      .Bounds(0).lLbound = 0
      .Bounds(0).cElements = lHeight
      .Bounds(1).lLbound = 0
      .Bounds(1).cElements = cDib.BytesPerScanLine()
      .pvData = cDib.DIBSectionBitsPtr
   End With
   CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4
   
   xOut = 0
   For x = 0 To cDib.BytesPerScanLine() - 4 Step 4
      If (lBit = 8) Then
         lBit = 0
         xOut = xOut + 1
      End If
      For y = 0 To lHeight - 1
         yOut = y
         If (bDib(x + 3, y) = 0) Then
            ' Output = 1
            b(xOut, yOut) = BitSet(b(xOut, yOut), lBit)
         Else
            ' Output = 0
         End If
      Next y
      lBit = lBit + 1
   Next x
   
   ' Clear the temporary array descriptor
   ' (This does not appear to be necessary, but
   ' for safety do it anyway)
   CopyMemory ByVal VarPtrArray(bDib), 0&, 4
   
End Sub
Private Function BitSet(ByVal b As Byte, ByVal lBit As Long) As Byte
   Select Case lBit
   Case 0
      b = b Or &H1
   Case 1
      b = b Or &H2
   Case 2
      b = b Or &H4
   Case 3
      b = b Or &H8
   Case 4
      b = b Or &H10
   Case 5
      b = b Or &H20
   Case 6
      b = b Or &H40
   Case 7
      b = b Or &H80
   End Select
   BitSet = b
End Function

Private Sub createIcon(cFI As cFileIcon)
Dim lIndex As Long
Dim i As Long
Dim iPos As Long
Dim lWidth As Long
Dim lHeight As Long
Dim sWidthHeight As String
   If (Option1.Value = True) Then
      lIndex = cFI.IconIndex(16, 16, 32)
      If (lIndex = 0) Then
         lIndex = cFI.AddImage(16, 16, 32)
      End If
      createIconAtSize cFI, lIndex, 16, 16
   End If
   If (Option2.Value = True) Then
      lIndex = cFI.IconIndex(32, 32, 32)
      If (lIndex = 0) Then
         lIndex = cFI.AddImage(32, 32, 32)
      End If
      createIconAtSize cFI, lIndex, 32, 32
   End If
   If (Option3.Value = True) Then
      lIndex = cFI.IconIndex(48, 48, 32)
      If (lIndex = 0) Then
         lIndex = cFI.AddImage(48, 48, 32)
      End If
      createIconAtSize cFI, lIndex, 48, 48
   End If
   If (Option4.Value = True) Then
      lIndex = cFI.IconIndex(64, 64, 32)
      If (lIndex = 0) Then
         lIndex = cFI.AddImage(64, 64, 32)
      End If
      createIconAtSize cFI, lIndex, 64, 64
   End If
      If (Option5.Value = True) Then
      lIndex = cFI.IconIndex(128, 128, 32)
      If (lIndex = 0) Then
         lIndex = cFI.AddImage(128, 128, 32)
      End If
      createIconAtSize cFI, lIndex, 128, 128
   End If
         If (Option6.Value = True) Then
      lIndex = cFI.IconIndex(24, 24, 32)
      If (lIndex = 0) Then
         lIndex = cFI.AddImage(24, 24, 32)
      End If
      createIconAtSize cFI, lIndex, 24, 24
   End If
End Sub

Private Sub openImage()
   
   Set m_cSource = New cAlphaDIBSection
   m_cSource.CreateFromPicture m_pic
   m_cSource.SetAlpha 255

End Sub

Private Sub setTransparentColour(ByVal lColor As Long)
   
   m_cSource.SetColourTransparent lColor
   
   m_lTransparentColour = lColor

   renderImage
   
End Sub

Private Sub renderImage()
   picsource.Cls
   m_cSource.AlphaPaintPicture picsource.hdc, _
      0, 0, _
      picsource.ScaleWidth \ Screen.TwipsPerPixelX, _
      picsource.ScaleHeight \ Screen.TwipsPerPixelY, _
      0, 0, _
      m_cSource.Width, m_cSource.Height
   picsource.Refresh
End Sub

Private Function fileExists(ByVal sFile As String) As Boolean
Dim sDir As String
   On Error Resume Next
   sDir = Dir(sFile)
   fileExists = (Len(sDir) > 0) And (Err.Number = 0)
End Function
Private Function GetCPUUsage() As Long
    
    Dim sbSysBasicInfo As SYSTEM_BASIC_INFORMATION
    Dim spSysPerforfInfo As SYSTEM_PERFORMANCE_INFORMATION
    Dim stSysTimeInfo As SYSTEM_TIME_INFORMATION
    Dim curIdle As Currency
    Dim curSystem As Currency
    Dim lngResult As Long
    
    GetCPUUsage = -1
    
    lngResult = NtQuerySystemInformation(SYSTEM_BASICINFORMATION, VarPtr(sbSysBasicInfo), LenB(sbSysBasicInfo), 0&)
    If lngResult <> NO_ERROR Then Exit Function
    
    lngResult = NtQuerySystemInformation(SYSTEM_TIMEINFORMATION, VarPtr(stSysTimeInfo), LenB(stSysTimeInfo), 0&)
    If lngResult <> NO_ERROR Then Exit Function
    
    lngResult = NtQuerySystemInformation(SYSTEM_PERFORMANCEINFORMATION, VarPtr(spSysPerforfInfo), LenB(spSysPerforfInfo), ByVal 0&)
    If lngResult <> NO_ERROR Then Exit Function
    curIdle = ConvertLI(spSysPerforfInfo.liIdleTime) - ConvertLI(lidOldIdle)
    curSystem = ConvertLI(stSysTimeInfo.liKeSystemTime) - ConvertLI(liOldSystem)
    If curSystem <> 0 Then curIdle = curIdle / curSystem
    curIdle = 100 - curIdle * 100 / sbSysBasicInfo.bKeNumberProcessors + 0.5
    GetCPUUsage = Int(curIdle)
    
    lidOldIdle = spSysPerforfInfo.liIdleTime
    liOldSystem = stSysTimeInfo.liKeSystemTime
End Function

Private Function ConvertLI(liToConvert As LARGE_INTEGER) As Currency
    CopyMemory ConvertLI, liToConvert, LenB(liToConvert)
End Function
Private Function GetErrorDescription(ByVal lErr As Long) As String
    Dim sReturn As String
    sReturn = String$(256, 32)
    FormatMessage FORMAT_MESSAGE_FROM_SYSTEM Or _
        FORMAT_MESSAGE_IGNORE_INSERTS, 0&, lErr, _
        0&, sReturn, Len(sReturn), ByVal 0
    sReturn = Trim(sReturn)
    GetErrorDescription = sReturn
End Function
Private Function GetProcessID(lpszProcessName As String) As Long
'RETUREN VALUES
'VALUE=-25 : FUNCTION FAILED
'VALUE<>-25 : SUCCEED
Dim pid    As Long
Dim pname As String
Dim a As String
a = Trim(LCase(lpszProcessName))
Dim my    As PROCESSENTRY32
Dim L    As Long
Dim l1    As Long
Dim flag    As Boolean
Dim mName    As String
Dim i    As Integer
L = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If L Then
    my.dwSize = 1060
End If
If (Process32First(L, my)) Then
    Do
        i = InStr(1, my.szExeFile, Chr(0))
        mName = LCase(Left(my.szExeFile, i - 1))
        If mName = a Then
            pid = my.th32ProcessID
            GetProcessID = pid
            Exit Function
        End If
Loop Until (Process32Next(L, my) < 1)
GetProcessID = -25
End If
End Function
Private Function GetProcessInfo(lpszProcessName As String, lpProcessInfo As PROCESSENTRY32) As Long
Dim pid    As Long
Dim pname As String
Dim a As String
a = Trim(LCase(lpszProcessName))
Dim my    As PROCESSENTRY32
Dim L    As Long
Dim l1    As Long
Dim flag    As Boolean
Dim mName    As String
Dim i    As Integer
L = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If L Then
    my.dwSize = 1060
End If
If (Process32First(L, my)) Then
    Do
        i = InStr(1, my.szExeFile, Chr(0))
        mName = LCase(Left(my.szExeFile, i - 1))
        If mName = a Then
            pid = my.th32ProcessID
            lpProcessInfo = my
            GetProcessInfo = 245
            Exit Function
        End If
Loop Until (Process32Next(L, my) < 1)
GetProcessInfo = -245
End If
End Function
Private Sub CloseScreenA(ByVal sWitch As Boolean)
If sWitch = True Then
CloseScreenFun GetForegroundWindow, WM_SYSCOMMAND, SC_MONITORPOWER, 1&
Else
CloseScreenFun GetForegroundWindow, WM_SYSCOMMAND, SC_MONITORPOWER, -1&
End If
End Sub
Public Function GetFolderName(hwnd As Long, Text As String) As String
On Error Resume Next
Dim bi As BROWSEINFO
Dim pidl As Long
Dim path As String
With bi
.hOwner = hwnd
.pidlRoot = 0&
.lpszTitle = Text
.ulFlags = BIF_NONEWFOLDERBUTTON
End With
pidl = SHBrowseForFolder(bi)
path = Space$(512)
If SHGetPathFromIDList(ByVal pidl, ByVal path) Then
GetFolderName = Left(path, InStr(path, Chr(0)) - 1)
End If
End Function
Sub GetProcessName(ByVal processID As Long, szExeName As String, szPathName As String)
On Error Resume Next
Dim my As PROCESSENTRY32
Dim hProcessHandle As Long
Dim success As Long
Dim L As Long
L = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
If L Then
my.dwSize = 1060
If (Process32First(L, my)) Then
Do
If my.th32ProcessID = processID Then
CloseHandle L
szExeName = Left$(my.szExeFile, InStr(1, my.szExeFile, Chr$(0)) - 1)
For L = Len(szExeName) To 1 Step -1
If Mid$(szExeName, L, 1) = "\" Then
Exit For
End If
Next L
szPathName = Left$(szExeName, L)
Exit Sub
End If
Loop Until (Process32Next(L, my) < 1)
End If
CloseHandle L
End If
End Sub
Private Sub CreateFile(lpPath As String, lpSize As Long)
On Error Resume Next
End Sub
Private Sub DisableClose(hwnd As Long, Optional ByVal MDIChild As Boolean)
On Error Resume Next
Exit Sub
Dim hSysMenu As Long
Dim nCnt As Long
Dim cID As Long
hSysMenu = GetSystemMenu(hwnd, False)
If hSysMenu = 0 Then
Exit Sub
End If
nCnt = GetMenuItemCount(hSysMenu)
If MDIChild Then
cID = 3
Else
cID = 1
End If
If nCnt Then
RemoveMenu hSysMenu, nCnt - cID, MF_BYPOSITION Or MF_REMOVE
RemoveMenu hSysMenu, nCnt - cID - 1, MF_BYPOSITION Or MF_REMOVE
DrawMenuBar hwnd
End If
End Sub
Private Function GetPassword(hwnd As Long) As String
On Error Resume Next
lpLength = SendMessage(hwnd, WM_GETTEXTLENGTH, 0, 0)
If lpLength > 0 Then
ReDim lpArray(lpLength + 1) As Byte
ReDim lpArray2(lpLength - 1) As Byte
CopyMemory lpArray(0), lpLength, 2
SendMessage hwnd, WM_GETTEXT, lpLength + 1, lpArray(0)
CopyMemory lpArray2(0), lpArray(0), lpLength
GetPassword = StrConv(lpArray2, vbUnicode)
Else
GetPassword = ""
End If
End Function
Private Function GetWindowClassName(hwnd As Long) As String
On Error Resume Next
Dim lpszWindowClassName As String * 256
lpszWindowClassName = Space(256)
GetClassName hwnd, lpszWindowClassName, 256
lpszWindowClassName = Trim(lpszWindowClassName)
GetWindowClassName = lpszWindowClassName
End Function
Private Sub AdjustToken()
On Error Resume Next
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_QUERY = &H8
Const SE_PRIVILEGE_ENABLED = &H2
Dim hdlProcessHandle As Long
Dim hdlTokenHandle As Long
Dim tmpLuid As LUID
Dim tkp As TOKEN_PRIVILEGES
Dim tkpNewButIgnored As TOKEN_PRIVILEGES
Dim lBufferNeeded As Long
hdlProcessHandle = GetCurrentProcess()
OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
TOKEN_QUERY), hdlTokenHandle
LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
tkp.PrivilegeCount = 1
tkp.TheLuid = tmpLuid
tkp.Attributes = SE_PRIVILEGE_ENABLED
AdjustTokenPrivileges hdlTokenHandle, False, _
tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
End Sub
Private Function HexOpen(lpFilePath As String, bSafe As Boolean) As String
Dim strFileName As String
Dim arr() As Byte
strFileName = App.path & "\2.jpg"
Open lpFilePath For Binary As #1
ReDim arr(LOF(1))
Get #1, , arr()
Close #1
Dim T As String
Dim L As Integer
Dim te As String
Dim ASCII As String
L = 0
T = ""
te = ""
ASCII = ""
Dim i
For i = LBound(arr) To UBound(arr)
te = UCase(Hex$(arr(i)))
If arr(i) >= 32 And arr(i) <= 126 Then
ASCII = ASCII & Chr(arr(i))
Else
ASCII = ASCII & "."
End If
If Len(te) = 1 Then te = "0" & te
T = T & te & " "
L = L + 1
If L = 16 Then
T = T & " "
ASCII = ASCII & " "
End If
If L = 32 Then
't = t & " " & ASCII & vbCrLf
T = T
ASCII = ""
L = 0
End If
If bSafe = True Then
If Len(T) >= 72 Then
T = Left(T, 72)
Exit For
End If
End If
Next
HexOpen = T
End Function
Private Function OpenAsHexDocument(lpFile As String, lpHeadOnly As Boolean) As String
On Error Resume Next
Dim strFileName As String
Dim arr() As Byte
strFileName = lpFile
If 245 = 245 Then
Open strFileName For Binary As #1
ReDim arr(LOF(1))
Get #1, , arr()
Close #1
Dim T As String
Dim L As Integer
Dim te As String
Dim ASCII As String
L = 0
T = ""
te = ""
ASCII = ""
Dim i
For i = LBound(arr) To UBound(arr)
te = UCase(Hex$(arr(i)))
If arr(i) >= 32 And arr(i) <= 126 Then
ASCII = ASCII & Chr(arr(i))
Else
ASCII = ASCII & "."
End If
If Len(te) = 1 Then te = "0" & te
T = T & te & " "
If Len(T) >= 72 And lpHeadOnly = True Then
Exit For
End If
L = L + 1
If L = 16 Then
T = T & " "
ASCII = ASCII & " "
End If
If L = 32 Then
T = T
ASCII = ""
L = 0
End If
Next
End If
If lpHeadOnly = True Then
OpenAsHexDocument = Left(T, 72)
Else
OpenAsHexDocument = T
End If
End Function
Private Sub EnumProcess()
Dim SnapShot As Long
Dim NextProcess As Long
Dim PE As PROCESSENTRY32 '创建进程快照
SnapShot = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0) '如果队列不为空则搜索
If SnapShot <> -1 Then '设置进程结构长度
PE.dwSize = Len(PE) '获取首个进程
NextProcess = Process32First(SnapShot, PE)
Do While NextProcess '可对进程序做相应处理
'获取下一个
NextProcess = Process32Next(SnapShot, PE)
Loop '释放进程句柄 CloseHandle (SnapShot)
End If
End Sub
Private Sub cmdPickColour_Click()
On Error Resume Next
Dim lColor As Long
Dim cD As New cCommonDialogLite
   OleTranslateColor picTransparentColour.BackColor, 0, lColor
   If (cD.VBChooseColor(lColor, FullOpen:=True, Owner:=Me.hwnd)) Then
      picTransparentColour.BackColor = lColor
      openImage
      setTransparentColour lColor
   End If
End Sub
Private Sub Command1_Click()
On Error Resume Next
With CommonDialog1
.DialogTitle = "保存ICO圖標"
.Filter = "ICO圖標(*.ICO)|*.ICO"
.ShowModalWindow = True
.hWndCall = hwnd
Dim IsCanceled As Boolean
IsCanceled = .ShowSave
End With
If IsCanceled = False Then
Exit Sub
End If
If CommonDialog1.Filename <> "" Then
If Dir(CommonDialog1.Filename, vbHidden Or vbNormal Or vbReadOnly Or vbSystem) = "" Then
      Dim cFI As New cFileIcon
      If 25 = 245 Then
      If (fileExists(CommonDialog1.Filename)) Then
         cFI.LoadIcon CommonDialog1.Filename
      End If
      End If
      createIcon cFI
      If cFI.SaveIcon(CommonDialog1.Filename) Then
        Exit Sub
    Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
Else
Dim ans As Integer
ans = MsgBox("文件 " & CommonDialog1.Filename & " 已存在，是否替換為當前文件?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill CommonDialog1.Filename
      If 25 = 245 Then
      If (fileExists(CommonDialog1.Filename)) Then
         cFI.LoadIcon CommonDialog1.Filename
      End If
      End If
      createIcon cFI
      If cFI.SaveIcon(CommonDialog1.Filename) Then
        Exit Sub
    Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      Else
      Exit Sub
      End If
End If
End If
End Sub

Private Sub Command2_Click()
On Error Resume Next
Dim ans As Integer
ans = MsgBox("確定初始化嗎?", vbQuestion + vbYesNo, "Ask")
If ans = vbNo Then
Exit Sub
End If
With txtUrl
.Text = ""
.Locked = True
End With
Image1.Picture = LoadPicture()
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Option6.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
picTransparentColour.BackColor = RGB(255, 255, 255)
Image1.Picture = LoadPicture()
With Me.VScroll1
.Height = PicturePreview.Height
.Width = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.HScroll1
.Width = PicturePreview.Width
.Height = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.Image1
.Left = 0
.TOp = 0
.Picture = LoadPicture("")
.Width = 0
.Height = 0
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Option6.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
With Me.Command2
.Enabled = True
.Cancel = False
End With
With Me.Command1
.Enabled = False
.Default = False
End With
With Me
.Left = Screen.Width / 2 - .Width / 2
.TOp = Screen.Height / 2 - .Height / 2
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
With Me.Option1
.Enabled = False
End With
With Me.Option2
.Enabled = False
End With
With Me.Option3
.Enabled = False
End With
With Me.Option4
.Enabled = False
End With
With Me.Option5
.Enabled = False
End With
With Me.Frame3
.Enabled = False
End With
Command1.Enabled = False
End Sub
Private Sub CommandBrowse_Click()
On Error GoTo ep
With CommonDialog1
.CancelError = True
.DialogTitle = "选择要装载的图片文件"
.Filter = "Pictures(*.jpg;*.jpeg;*.bmp;*.bip;*.gif;*.wmf)|*.jpg;*.jpeg;*.bmp;*.bip;*.gif;*.wmf"
.ShowModalWindow = True
.hWndCall = hwnd
.CancelError = True
Dim IsCanceled As Boolean
IsCanceled = .ShowOpen
End With
If IsCanceled = False Then
Exit Sub
End If
If Trim(CommonDialog1.Filename) <> "" Then
If Image1.Picture = IC.Picture Then
Image1.Picture = LoadPicture(CommonDialog1.Filename)
picsource.Picture = LoadPicture(CommonDialog1.Filename)
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(CommonDialog1.Filename)
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      setTransparentColour picTransparentColour.BackColor
txtUrl.Text = CommonDialog1.Filename
txtUrl.Locked = True
PicturePreview.Enabled = True
Image1.Enabled = True
Me.Image1.Left = 0
Me.Image1.TOp = 0
Me.HScroll1.Enabled = True
Me.VScroll1.Enabled = True
If Me.Image1.Height <= Me.PicturePreview.Height And Me.Image1.Width <= Me.PicturePreview.Width Then
Me.Image1.Left = Me.PicturePreview.Width / 2 - Me.Image1.Width / 2
Me.Image1.TOp = Me.PicturePreview.Height / 2 - Me.Image1.Height / 2
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = False
Option5.Enabled = True
Option4.Enabled = True
Option3.Enabled = True
Option2.Enabled = True
Option1.Enabled = True
Option6.Enabled = True
Frame3.Enabled = True
Option2.Value = True
Command1.Enabled = True
cmdPickColour.Enabled = True
picTransparentColour.Enabled = True
lblTransparentColour.Enabled = True
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 32, 32
.Refresh
End With
With Me.Command1
.Enabled = True
.Default = True
End With
With Me.Command2
.Enabled = True
.Cancel = True
End With
Exit Sub
End If
If Me.Image1.Width <= Me.PicturePreview.Width Then
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = True
Me.Image1.Left = Me.PicturePreview.Width / 2 - Me.Image1.Width / 2
End If
If Me.Image1.Height <= Me.PicturePreview.Height Then
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = True
Me.Image1.TOp = Me.PicturePreview.Height / 2 - Me.Image1.Height / 2
End If
Me.VScroll1.Max = Me.Image1.Height - Me.PicturePreview.Height
Me.HScroll1.Max = Me.Image1.Width - Me.PicturePreview.Width
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = Image1.Width / 10
Me.HScroll1.SmallChange = Image1.Width / 50
Me.VScroll1.LargeChange = Image1.Width / 10
Me.VScroll1.SmallChange = Image1.Width / 50
Option5.Enabled = True
Option4.Enabled = True
Option3.Enabled = True
Option2.Enabled = True
Option1.Enabled = True
Option6.Enabled = True
Frame3.Enabled = True
Option2.Value = True
Command1.Enabled = True
cmdPickColour.Enabled = True
picTransparentColour.Enabled = True
lblTransparentColour.Enabled = True
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 32, 32
.Refresh
End With
With Me.Command1
.Enabled = True
.Default = True
End With
With Me.Command2
.Enabled = True
.Cancel = True
End With
Else
Dim ans As Integer
ans = MsgBox("已經載入了圖片，是否替換為當前選定圖片?", vbQuestion + vbYesNo, "Ask")
If ans = vbNo Then
Exit Sub
End If
Image1.Picture = LoadPicture(CommonDialog1.Filename)
picsource.Picture = LoadPicture(CommonDialog1.Filename)
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(CommonDialog1.Filename)
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      setTransparentColour picTransparentColour.BackColor
txtUrl.Text = CommonDialog1.Filename
txtUrl.Locked = True
PicturePreview.Enabled = True
Image1.Enabled = True
Me.Image1.Left = 0
Me.Image1.TOp = 0
Me.HScroll1.Enabled = True
Me.VScroll1.Enabled = True
If Me.Image1.Height <= Me.PicturePreview.Height And Me.Image1.Width <= Me.PicturePreview.Width Then
Me.Image1.Left = Me.PicturePreview.Width / 2 - Me.Image1.Width / 2
Me.Image1.TOp = Me.PicturePreview.Height / 2 - Me.Image1.Height / 2
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = False
Option5.Enabled = True
Option4.Enabled = True
Option3.Enabled = True
Option2.Enabled = True
Option1.Enabled = True
Option6.Enabled = True
Frame3.Enabled = True
Option2.Value = True
Command1.Enabled = True
cmdPickColour.Enabled = True
picTransparentColour.Enabled = True
lblTransparentColour.Enabled = True
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 32, 32
.Refresh
End With
With Me.Command1
.Enabled = True
.Default = True
End With
With Me.Command2
.Enabled = True
.Cancel = True
End With
Exit Sub
End If
If Me.Image1.Width <= Me.PicturePreview.Width Then
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = True
Me.Image1.Left = Me.PicturePreview.Width / 2 - Me.Image1.Width / 2
End If
If Me.Image1.Height <= Me.PicturePreview.Height Then
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = True
Me.Image1.TOp = Me.PicturePreview.Height / 2 - Me.Image1.Height / 2
End If
Me.VScroll1.Max = Me.Image1.Height - Me.PicturePreview.Height
Me.HScroll1.Max = Me.Image1.Width - Me.PicturePreview.Width
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = Image1.Width / 10
Me.HScroll1.SmallChange = Image1.Width / 50
Me.VScroll1.LargeChange = Image1.Width / 10
Me.VScroll1.SmallChange = Image1.Width / 50
Option5.Enabled = True
Option4.Enabled = True
Option3.Enabled = True
Option2.Enabled = True
Option1.Enabled = True
Option6.Enabled = True
Frame3.Enabled = True
Option2.Value = True
Command1.Enabled = True
cmdPickColour.Enabled = True
picTransparentColour.Enabled = True
lblTransparentColour.Enabled = True
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 32, 32
.Refresh
End With
With Me.Command1
.Enabled = True
.Default = True
End With
With Me.Command2
.Enabled = True
.Cancel = True
End With
End If
End If
Exit Sub
ep:
MsgBox "發生内部錯誤：" & vbCrLf & Err.Description, vbCritical, "Error"
With txtUrl
.Text = ""
.Locked = True
End With
Image1.Picture = LoadPicture()
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Option6.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
picTransparentColour.BackColor = RGB(255, 255, 255)
Image1.Picture = LoadPicture()
With Me.VScroll1
.Height = PicturePreview.Height
.Width = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.HScroll1
.Width = PicturePreview.Width
.Height = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.Image1
.Left = 0
.TOp = 0
.Picture = LoadPicture("")
.Width = 0
.Height = 0
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Option6.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
With Me.Command2
.Enabled = True
.Cancel = False
End With
With Me.Command1
.Enabled = False
.Default = False
End With
With Me
.Left = Screen.Width / 2 - .Width / 2
.TOp = Screen.Height / 2 - .Height / 2
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
With Me.Option1
.Enabled = False
End With
With Me.Option2
.Enabled = False
End With
With Me.Option3
.Enabled = False
End With
With Me.Option4
.Enabled = False
End With
With Me.Option5
.Enabled = False
End With
With Me.Frame3
.Enabled = False
End With
Command1.Enabled = False
On Error Resume Next
With txtUrl
.Text = ""
.Locked = True
End With
Image1.Picture = LoadPicture()
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Option6.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
Image1.Picture = LoadPicture()
With Me.VScroll1
.Height = PicturePreview.Height
.Width = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.HScroll1
.Width = PicturePreview.Width
.Height = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.Image1
.Left = 0
.TOp = 0
.Picture = LoadPicture("")
.Width = 0
.Height = 0
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Option6.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
With Me.Command2
.Enabled = True
.Cancel = False
End With
With Me.Command1
.Enabled = False
.Default = False
End With
With Me
.Left = Screen.Width / 2 - .Width / 2
.TOp = Screen.Height / 2 - .Height / 2
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
With Me.Option1
.Enabled = False
End With
With Me.Option2
.Enabled = False
End With
With Me.Option3
.Enabled = False
End With
With Me.Option4
.Enabled = False
End With
With Me.Option5
.Enabled = False
End With
With Me.Frame3
.Enabled = False
End With
Command1.Enabled = False
End Sub
Private Sub Form_Initialize()
On Error Resume Next
   HookAttach
   
   m_lTransparentColour = CLR_INVALID
End Sub
Private Sub Form_Load()
On Error Resume Next
HookDetach
With txtUrl
.Text = ""
.Locked = True
End With
Image1.Picture = LoadPicture()
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Option6.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
picTransparentColour.BackColor = RGB(255, 255, 255)
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
Image1.Picture = LoadPicture()
With Me.VScroll1
.Height = PicturePreview.Height
.Width = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.HScroll1
.Width = PicturePreview.Width
.Height = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.Image1
.Left = 0
.TOp = 0
.Picture = LoadPicture("")
.Width = 0
.Height = 0
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Option6.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
With Me.Command2
.Enabled = True
.Cancel = False
End With
With Me.Command1
.Enabled = False
.Default = False
End With
With Me
.Left = Screen.Width / 2 - .Width / 2
.TOp = Screen.Height / 2 - .Height / 2
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
With Me.Option1
.Enabled = False
End With
With Me.Option2
.Enabled = False
End With
With Me.Option3
.Enabled = False
End With
With Me.Option4
.Enabled = False
End With
With Me.Option5
.Enabled = False
End With
With Me.Frame3
.Enabled = False
End With
Command1.Enabled = False
End Sub
Private Sub mnuAbout_Click()
On Error Resume Next
frmAbout.Show 1
End Sub
Private Sub mnuExit_Click()
On Error Resume Next
Dim ans As Integer
ans = MsgBox("確定要退出嗎?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Unload Me
Close
End
Else
Exit Sub
End If
End Sub
Private Sub mnuGet_Click()
On Error Resume Next
Form1.Show 1
End Sub
Private Sub mnuInit_Click()
On Error Resume Next
Dim ans As Integer
ans = MsgBox("確定初始化嗎?", vbQuestion + vbYesNo, "Ask")
If ans = vbNo Then
Exit Sub
End If
With txtUrl
.Text = ""
.Locked = True
End With
Image1.Picture = LoadPicture()
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Option6.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
picTransparentColour.BackColor = RGB(255, 255, 255)
Image1.Picture = LoadPicture()
With Me.VScroll1
.Height = PicturePreview.Height
.Width = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.HScroll1
.Width = PicturePreview.Width
.Height = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.Image1
.Left = 0
.TOp = 0
.Picture = LoadPicture("")
.Width = 0
.Height = 0
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Option6.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
With Me.Command2
.Enabled = True
.Cancel = False
End With
With Me.Command1
.Enabled = False
.Default = False
End With
With Me
.Left = Screen.Width / 2 - .Width / 2
.TOp = Screen.Height / 2 - .Height / 2
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
With Me.Option1
.Enabled = False
End With
With Me.Option2
.Enabled = False
End With
With Me.Option3
.Enabled = False
End With
With Me.Option4
.Enabled = False
End With
With Me.Option5
.Enabled = False
End With
With Me.Frame3
.Enabled = False
End With
Command1.Enabled = False
With txtUrl
.Text = ""
.Locked = True
End With
Image1.Picture = LoadPicture()
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Option6.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
picTransparentColour.BackColor = RGB(255, 255, 255)
Image1.Picture = LoadPicture()
With Me.VScroll1
.Height = PicturePreview.Height
.Width = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.HScroll1
.Width = PicturePreview.Width
.Height = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.Image1
.Left = 0
.TOp = 0
.Picture = LoadPicture("")
.Width = 0
.Height = 0
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Option6.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
With Me.Command2
.Enabled = True
.Cancel = False
End With
With Me.Command1
.Enabled = False
.Default = False
End With
With Me
.Left = Screen.Width / 2 - .Width / 2
.TOp = Screen.Height / 2 - .Height / 2
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
With Me.Option1
.Enabled = False
End With
With Me.Option2
.Enabled = False
End With
With Me.Option3
.Enabled = False
End With
With Me.Option4
.Enabled = False
End With
With Me.Option5
.Enabled = False
End With
With Me.Frame3
.Enabled = False
End With
Command1.Enabled = False
End Sub
Private Sub mnuOpen_Click()
On Error GoTo ep
With CommonDialog1
.CancelError = True
.DialogTitle = "选择要装载的图片文件"
.Filter = "Pictures(*.jpg;*.jpeg;*.bmp;*.bip;*.gif;*.wmf)|*.jpg;*.jpeg;*.bmp;*.bip;*.gif;*.wmf"
.ShowModalWindow = True
.hWndCall = hwnd
.CancelError = True
Dim IsCanceled As Boolean
IsCanceled = .ShowOpen
End With
If IsCanceled = False Then
Exit Sub
End If
If Trim(CommonDialog1.Filename) <> "" Then
If Image1.Picture = IC.Picture Then
Image1.Picture = LoadPicture(CommonDialog1.Filename)
picsource.Picture = LoadPicture(CommonDialog1.Filename)
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(CommonDialog1.Filename)
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      setTransparentColour picTransparentColour.BackColor
txtUrl.Text = CommonDialog1.Filename
txtUrl.Locked = True
PicturePreview.Enabled = True
Image1.Enabled = True
Me.Image1.Left = 0
Me.Image1.TOp = 0
Me.HScroll1.Enabled = True
Me.VScroll1.Enabled = True
If Me.Image1.Height <= Me.PicturePreview.Height And Me.Image1.Width <= Me.PicturePreview.Width Then
Me.Image1.Left = Me.PicturePreview.Width / 2 - Me.Image1.Width / 2
Me.Image1.TOp = Me.PicturePreview.Height / 2 - Me.Image1.Height / 2
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = False
Option5.Enabled = True
Option4.Enabled = True
Option3.Enabled = True
Option2.Enabled = True
Option1.Enabled = True
Option6.Enabled = True
Frame3.Enabled = True
Option2.Value = True
Command1.Enabled = True
cmdPickColour.Enabled = True
picTransparentColour.Enabled = True
lblTransparentColour.Enabled = True
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 32, 32
.Refresh
End With
With Me.Command1
.Enabled = True
.Default = True
End With
With Me.Command2
.Enabled = True
.Cancel = True
End With
Exit Sub
End If
If Me.Image1.Width <= Me.PicturePreview.Width Then
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = True
Me.Image1.Left = Me.PicturePreview.Width / 2 - Me.Image1.Width / 2
End If
If Me.Image1.Height <= Me.PicturePreview.Height Then
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = True
Me.Image1.TOp = Me.PicturePreview.Height / 2 - Me.Image1.Height / 2
End If
Me.VScroll1.Max = Me.Image1.Height - Me.PicturePreview.Height
Me.HScroll1.Max = Me.Image1.Width - Me.PicturePreview.Width
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = Image1.Width / 10
Me.HScroll1.SmallChange = Image1.Width / 50
Me.VScroll1.LargeChange = Image1.Width / 10
Me.VScroll1.SmallChange = Image1.Width / 50
Option5.Enabled = True
Option4.Enabled = True
Option3.Enabled = True
Option2.Enabled = True
Option1.Enabled = True
Option6.Enabled = True
Frame3.Enabled = True
Option2.Value = True
Command1.Enabled = True
cmdPickColour.Enabled = True
picTransparentColour.Enabled = True
lblTransparentColour.Enabled = True
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 32, 32
.Refresh
End With
With Me.Command1
.Enabled = True
.Default = True
End With
With Me.Command2
.Enabled = True
.Cancel = True
End With
Else
Dim ans As Integer
ans = MsgBox("已經載入了圖片，是否替換為當前選定圖片?", vbQuestion + vbYesNo, "Ask")
If ans = vbNo Then
Exit Sub
End If
Image1.Picture = LoadPicture(CommonDialog1.Filename)
picsource.Picture = LoadPicture(CommonDialog1.Filename)
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(CommonDialog1.Filename)
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      setTransparentColour picTransparentColour.BackColor
txtUrl.Text = CommonDialog1.Filename
txtUrl.Locked = True
PicturePreview.Enabled = True
Image1.Enabled = True
Me.Image1.Left = 0
Me.Image1.TOp = 0
Me.HScroll1.Enabled = True
Me.VScroll1.Enabled = True
If Me.Image1.Height <= Me.PicturePreview.Height And Me.Image1.Width <= Me.PicturePreview.Width Then
Me.Image1.Left = Me.PicturePreview.Width / 2 - Me.Image1.Width / 2
Me.Image1.TOp = Me.PicturePreview.Height / 2 - Me.Image1.Height / 2
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = False
Option5.Enabled = True
Option4.Enabled = True
Option3.Enabled = True
Option2.Enabled = True
Option1.Enabled = True
Option6.Enabled = True
Frame3.Enabled = True
Option2.Value = True
Command1.Enabled = True
cmdPickColour.Enabled = True
picTransparentColour.Enabled = True
lblTransparentColour.Enabled = True
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 32, 32
.Refresh
End With
With Me.Command1
.Enabled = True
.Default = True
End With
With Me.Command2
.Enabled = True
.Cancel = True
End With
Exit Sub
End If
If Me.Image1.Width <= Me.PicturePreview.Width Then
Me.HScroll1.Enabled = False
Me.VScroll1.Enabled = True
Me.Image1.Left = Me.PicturePreview.Width / 2 - Me.Image1.Width / 2
End If
If Me.Image1.Height <= Me.PicturePreview.Height Then
Me.VScroll1.Enabled = False
Me.HScroll1.Enabled = True
Me.Image1.TOp = Me.PicturePreview.Height / 2 - Me.Image1.Height / 2
End If
Me.VScroll1.Max = Me.Image1.Height - Me.PicturePreview.Height
Me.HScroll1.Max = Me.Image1.Width - Me.PicturePreview.Width
Me.VScroll1.Value = 0
Me.HScroll1.Value = 0
Me.HScroll1.LargeChange = Image1.Width / 10
Me.HScroll1.SmallChange = Image1.Width / 50
Me.VScroll1.LargeChange = Image1.Width / 10
Me.VScroll1.SmallChange = Image1.Width / 50
Option5.Enabled = True
Option4.Enabled = True
Option3.Enabled = True
Option2.Enabled = True
Option1.Enabled = True
Option6.Enabled = True
Frame3.Enabled = True
Option2.Value = True
Command1.Enabled = True
cmdPickColour.Enabled = True
picTransparentColour.Enabled = True
lblTransparentColour.Enabled = True
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 32, 32
.Refresh
End With
With Me.Command1
.Enabled = True
.Default = True
End With
With Me.Command2
.Enabled = True
.Cancel = True
End With
End If
End If
Exit Sub
ep:
MsgBox "發生内部錯誤：" & vbCrLf & Err.Description, vbCritical, "Error"
With txtUrl
.Text = ""
.Locked = True
End With
Image1.Picture = LoadPicture()
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Option6.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
picTransparentColour.BackColor = RGB(255, 255, 255)
Image1.Picture = LoadPicture()
With Me.VScroll1
.Height = PicturePreview.Height
.Width = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.HScroll1
.Width = PicturePreview.Width
.Height = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.Image1
.Left = 0
.TOp = 0
.Picture = LoadPicture("")
.Width = 0
.Height = 0
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Option6.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
With Me.Command2
.Enabled = True
.Cancel = False
End With
With Me.Command1
.Enabled = False
.Default = False
End With
With Me
.Left = Screen.Width / 2 - .Width / 2
.TOp = Screen.Height / 2 - .Height / 2
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
With Me.Option1
.Enabled = False
End With
With Me.Option2
.Enabled = False
End With
With Me.Option3
.Enabled = False
End With
With Me.Option4
.Enabled = False
End With
With Me.Option5
.Enabled = False
End With
With Me.Frame3
.Enabled = False
End With
Command1.Enabled = False
On Error Resume Next
With txtUrl
.Text = ""
.Locked = True
End With
Image1.Picture = LoadPicture()
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Option6.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
Image1.Picture = LoadPicture()
With Me.VScroll1
.Height = PicturePreview.Height
.Width = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.HScroll1
.Width = PicturePreview.Width
.Height = 255
.Enabled = False
.Min = 0
.Max = 0
.Value = 0
End With
With Me.Image1
.Left = 0
.TOp = 0
.Picture = LoadPicture("")
.Width = 0
.Height = 0
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Option6.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
cmdPickColour.Enabled = False
picTransparentColour.Enabled = False
lblTransparentColour.Enabled = False
With Me.Command2
.Enabled = True
.Cancel = False
End With
With Me.Command1
.Enabled = False
.Default = False
End With
With Me
.Left = Screen.Width / 2 - .Width / 2
.TOp = Screen.Height / 2 - .Height / 2
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
With Me.Option1
.Enabled = False
End With
With Me.Option2
.Enabled = False
End With
With Me.Option3
.Enabled = False
End With
With Me.Option4
.Enabled = False
End With
With Me.Option5
.Enabled = False
End With
With Me.Frame3
.Enabled = False
End With
Command1.Enabled = False
End Sub
Private Sub mnuWizard_Click()
On Error Resume Next
frmAlphaIconCreator.Show 1
End Sub
Private Sub Option6_Click()
On Error Resume Next
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 28
.Width = 28
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 24, 24
.Refresh
End With
If 25 = 245 Then
With Me.picsource
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 20
.Width = 20
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 16, 16
.Refresh
.Picture = Image1.Picture
End With
End If
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub picTransparentColour_Click()
On Error Resume Next
Dim lColor As Long
Dim cD As New cCommonDialogLite
   OleTranslateColor picTransparentColour.BackColor, 0, lColor
   If (cD.VBChooseColor(lColor, FullOpen:=True, Owner:=Me.hwnd)) Then
      picTransparentColour.BackColor = lColor
      openImage
      setTransparentColour lColor
   End If
End Sub
Private Sub VScroll1_Change()
On Error Resume Next
Image1.TOp = -VScroll1.Value
If -VScroll1.Value > 0 Then
Image1.TOp = VScroll1.Value
End If
End Sub
Private Sub hscroll1_change()
On Error Resume Next
Image1.Left = -HScroll1.Value
If -HScroll1.Value > 0 Then
Image1.Left = HScroll1.Value
End If
End Sub
Private Sub Option1_Click()
On Error Resume Next
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 20
.Width = 20
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 16, 16
.Refresh
End With
If 25 = 245 Then
With Me.picsource
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 20
.Width = 20
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 16, 16
.Refresh
.Picture = Image1.Picture
End With
End If
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub Option2_Click()
On Error Resume Next
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 32, 32
.Refresh
End With
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub Option3_Click()
On Error Resume Next
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 52
.Width = 52
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 48, 48
.Refresh
End With
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub Option4_Click()
On Error Resume Next
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 68
.Width = 68
.Left = 0
.TOp = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 64, 64
.Refresh
End With
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub Option5_Click()
On Error Resume Next
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 132
.Width = 132
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.TOp = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 128, 128
.Refresh
End With
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
