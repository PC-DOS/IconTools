VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Icon Tools - PC-DOS Workshop"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   6000
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":08CA
   ScaleHeight     =   279
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   1980
      Left            =   3870
      ScaleHeight     =   132
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   132
      TabIndex        =   13
      Top             =   270
      Width           =   1980
   End
   Begin VB.Frame Frame2 
      Caption         =   "图标预览"
      Height          =   3855
      Left            =   3720
      TabIndex        =   12
      Top             =   30
      Width           =   2280
      Begin VB.Frame Frame3 
         Caption         =   "选项"
         Height          =   1365
         Left            =   60
         TabIndex        =   14
         Top             =   2400
         Width           =   2145
         Begin VB.ComboBox Combo1 
            Height          =   300
            ItemData        =   "Form1.frx":1794
            Left            =   120
            List            =   "Form1.frx":17AA
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   435
            Width           =   1965
         End
         Begin VB.ComboBox Combo2 
            Height          =   300
            ItemData        =   "Form1.frx":17EE
            Left            =   120
            List            =   "Form1.frx":1804
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   1005
            Width           =   1950
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "绘图大小:"
            Height          =   180
            Left            =   120
            TabIndex        =   18
            Top             =   210
            Width           =   810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "绘图背景:"
            Height          =   180
            Left            =   120
            TabIndex        =   17
            Top             =   780
            Width           =   810
         End
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   2145
         Left            =   75
         Top             =   165
         Width           =   2130
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         Index           =   1
         X1              =   75
         X2              =   2220
         Y1              =   2370
         Y2              =   2370
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         Index           =   0
         X1              =   75
         X2              =   2220
         Y1              =   2355
         Y2              =   2355
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "信息"
      Height          =   1830
      Left            =   30
      TabIndex        =   7
      Top             =   2055
      Width           =   3585
      Begin VB.TextBox Text1 
         Height          =   750
         Left            =   75
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Top             =   390
         Width           =   3405
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1440
         Picture         =   "Form1.frx":182C
         Top             =   420
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   75
         TabIndex        =   10
         Top             =   1410
         Width           =   3420
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "图标个数"
         Height          =   180
         Index           =   1
         Left            =   60
         TabIndex        =   9
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件"
         Height          =   180
         Index           =   0
         Left            =   60
         TabIndex        =   8
         Top             =   195
         Width           =   360
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   825
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   3555
      TabIndex        =   4
      Top             =   1185
      Width           =   3615
      Begin VB.CommandButton Command7 
         Caption         =   "下一个(&N)"
         Enabled         =   0   'False
         Height          =   735
         Left            =   2400
         Picture         =   "Form1.frx":20F6
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   15
         Width           =   1155
      End
      Begin VB.CommandButton Command6 
         Caption         =   "上一个(&L)"
         Enabled         =   0   'False
         Height          =   735
         Left            =   15
         Picture         =   "Form1.frx":2538
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   15
         Width           =   1155
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "图标指针位置:"
         Height          =   180
         Left            =   1185
         TabIndex        =   20
         Top             =   45
         Width           =   1170
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1170
         TabIndex        =   19
         Top             =   270
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   1155
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   3555
      TabIndex        =   0
      Top             =   0
      Width           =   3615
      Begin VB.CommandButton Command5 
         Cancel          =   -1  'True
         Caption         =   "退出(&E)"
         Height          =   1035
         Left            =   2460
         Picture         =   "Form1.frx":297A
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   30
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "保存(&S)"
         Enabled         =   0   'False
         Height          =   1035
         Left            =   1230
         Picture         =   "Form1.frx":2DBC
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   30
         Width           =   1110
      End
      Begin VB.CommandButton Command1 
         Caption         =   "打开文件(&O)"
         Height          =   1035
         Left            =   15
         Picture         =   "Form1.frx":31FE
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   30
         Width           =   1110
      End
   End
   Begin VB.Label StateBar1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "准备就绪"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   15
      TabIndex        =   21
      Top             =   3885
      Width           =   5985
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuOpen 
         Caption         =   "打开Win32PE文件(&O)..."
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSave 
         Caption         =   "另存为(&S)..."
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuClose 
         Caption         =   "关闭打开的文件(&C)"
         Shortcut        =   ^E
      End
      Begin VB.Menu bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuConvert 
         Caption         =   "图片转换器(&P)..."
      End
      Begin VB.Menu mnuCopy2 
         Caption         =   "复制图标(&C)"
      End
      Begin VB.Menu mnuPage 
         Caption         =   "页面设置(&A)"
         Begin VB.Menu mnuSizeLead 
            Caption         =   "绘图区域大小(&D)"
            Begin VB.Menu mnuSize 
               Caption         =   "(默认)"
               Index           =   0
            End
            Begin VB.Menu mnuSize 
               Caption         =   "16*16"
               Index           =   1
            End
            Begin VB.Menu mnuSize 
               Caption         =   "32*32"
               Index           =   2
            End
            Begin VB.Menu mnuSize 
               Caption         =   "48*48"
               Index           =   3
            End
            Begin VB.Menu mnuSize 
               Caption         =   "64*64"
               Index           =   4
            End
            Begin VB.Menu mnuSize 
               Caption         =   "128*128"
               Index           =   5
            End
         End
         Begin VB.Menu mnuBGCLead 
            Caption         =   "绘图区域背景色(&W)"
            Begin VB.Menu mnuBGC 
               Caption         =   "白色"
               Index           =   0
            End
            Begin VB.Menu mnuBGC 
               Caption         =   "蓝色"
               Index           =   1
            End
            Begin VB.Menu mnuBGC 
               Caption         =   "绿色"
               Index           =   2
            End
            Begin VB.Menu mnuBGC 
               Caption         =   "红色"
               Index           =   3
            End
            Begin VB.Menu mnuBGC 
               Caption         =   "黄色"
               Index           =   4
            End
            Begin VB.Menu mnuBGC 
               Caption         =   "黑色"
               Index           =   5
            End
         End
      End
      Begin VB.Menu bar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "退出(&E)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuCopy 
         Caption         =   "复制(&C)"
         Enabled         =   0   'False
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuSize2 
         Caption         =   "绘图区域大小"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuSizeView 
         Caption         =   "    (Default)"
         Index           =   0
      End
      Begin VB.Menu mnuSizeView 
         Caption         =   "    16*16"
         Index           =   1
      End
      Begin VB.Menu mnuSizeView 
         Caption         =   "    32*32"
         Index           =   2
      End
      Begin VB.Menu mnuSizeView 
         Caption         =   "    48*48"
         Index           =   3
      End
      Begin VB.Menu mnuSizeView 
         Caption         =   "    64*64"
         Index           =   4
      End
      Begin VB.Menu mnuSizeView 
         Caption         =   "    128*128"
         Index           =   5
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu mnuBack 
         Caption         =   "绘图区域背景"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuColor 
         Caption         =   "    白色"
         Index           =   0
      End
      Begin VB.Menu mnuColor 
         Caption         =   "    蓝色"
         Index           =   1
      End
      Begin VB.Menu mnuColor 
         Caption         =   "    绿色"
         Index           =   2
      End
      Begin VB.Menu mnuColor 
         Caption         =   "    红色"
         Index           =   3
      End
      Begin VB.Menu mnuColor 
         Caption         =   "    黄色"
         Index           =   4
      End
      Begin VB.Menu mnuColor 
         Caption         =   "    黑色"
         Index           =   5
      End
      Begin VB.Menu b245 
         Caption         =   "-"
      End
      Begin VB.Menu mnuJump 
         Caption         =   "跳转到(&J)..."
         Enabled         =   0   'False
         Shortcut        =   ^J
      End
   End
   Begin VB.Menu mnuCVS 
      Caption         =   "工具(&T)"
      Begin VB.Menu mnuRun 
         Caption         =   "将图像转换为图标(&R)..."
         Shortcut        =   ^R
      End
      Begin VB.Menu b8 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSaveExit 
         Caption         =   "退出时保存设置(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuShowCurrent 
         Caption         =   "显示保存的设置(&O)"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "清除保存的设置(&C)"
      End
      Begin VB.Menu mnuHide 
         Caption         =   "隐藏顶级菜单(&H)"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuAbout 
         Caption         =   "关于 Icon Tools(A)..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function ExtractIcon Lib "shell32.dll" Alias "ExtractIconA" (ByVal hInst As Long, ByVal lpszExeFileName As String, ByVal nIconIndex As Long) As Long
Private Declare Function DrawIconEx Lib "user32" (ByVal hdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Const DI_COMPACT = &H4
Const DI_DEFAULTSIZE = &H8
Const DI_IMAGE = &H2
Const DI_MASK = &H1
Const DI_NORMAL = &H3
Dim tmp As Long
Dim cur As Long
Dim hIco As Long
Dim cod As Boolean
Private Sub Combo1_Change()
On Error Resume Next
With Me.StateBar1
.Alignment = 0
.Caption = "页面设置更改"
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo2.ListIndex = 0 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbWhite
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 1 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbBlue
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 2 Then
Picture3.BackColor = vbGreen
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 3 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbRed
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 4 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbYellow
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 5 Then
Picture3.BackColor = vbBlack
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 0 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbWhite
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 1 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbBlue
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 2 Then
Picture3.BackColor = vbGreen
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 3 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbRed
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 4 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbYellow
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 5 Then
Picture3.BackColor = vbBlack
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
Dim cur1 As Integer
For cur1 = 0 To Me.mnuColor.UBound
Me.mnuColor(cur1).Checked = False
Next
Me.mnuColor(Combo2.ListIndex).Checked = True
For cur1 = 0 To Me.mnuBGC.UBound
mnuBGC(cur1).Checked = False
Next
Me.mnuBGC(Combo2.ListIndex).Checked = True
For cur1 = 0 To Me.mnuSize.UBound
Me.mnuSize(cur1).Checked = False
Next
Me.mnuSize(Combo1.ListIndex).Checked = True
For cur1 = 0 To Me.mnuSizeView.UBound
mnuSizeView(cur1).Checked = False
Next
Me.mnuSizeView(Combo1.ListIndex).Checked = True
Select Case Me.mnuSaveExit.Checked
Case True
SaveSetting "Icon Tools", "Options", "DrawSize", Me.Combo1.ListIndex
SaveSetting "Icon Tools", "Options", "BackColor", Me.Combo2.ListIndex
Case False
SaveSetting "Icon Tools", "Options", "DrawSize", 0
SaveSetting "Icon Tools", "Options", "BackColor", 0
End Select
End Sub
Private Sub Combo1_Click()
On Error Resume Next
With Me.StateBar1
.Alignment = 0
.Caption = "页面设置更改"
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo2.ListIndex = 0 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbWhite
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 1 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbBlue
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 2 Then
Picture3.BackColor = vbGreen
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 3 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbRed
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 4 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbYellow
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 5 Then
Picture3.BackColor = vbBlack
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 0 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbWhite
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 1 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbBlue
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 2 Then
Picture3.BackColor = vbGreen
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 3 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbRed
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 4 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbYellow
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 5 Then
Picture3.BackColor = vbBlack
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
Dim cur1 As Integer
For cur1 = 0 To Me.mnuColor.UBound
Me.mnuColor(cur1).Checked = False
Next
Me.mnuColor(Combo2.ListIndex).Checked = True
For cur1 = 0 To Me.mnuBGC.UBound
mnuBGC(cur1).Checked = False
Next
Me.mnuBGC(Combo2.ListIndex).Checked = True
For cur1 = 0 To Me.mnuSize.UBound
Me.mnuSize(cur1).Checked = False
Next
Me.mnuSize(Combo1.ListIndex).Checked = True
For cur1 = 0 To Me.mnuSizeView.UBound
mnuSizeView(cur1).Checked = False
Next
Me.mnuSizeView(Combo1.ListIndex).Checked = True
Select Case Me.mnuSaveExit.Checked
Case True
SaveSetting "Icon Tools", "Options", "DrawSize", Me.Combo1.ListIndex
SaveSetting "Icon Tools", "Options", "BackColor", Me.Combo2.ListIndex
Case False
SaveSetting "Icon Tools", "Options", "DrawSize", 0
SaveSetting "Icon Tools", "Options", "BackColor", 0
End Select
End Sub
Private Sub Combo2_Change()
On Error Resume Next
With Me.StateBar1
.Alignment = 0
.Caption = "页面设置更改"
End With
If Combo2.ListIndex = 0 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbWhite
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 1 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbBlue
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 2 Then
Picture3.BackColor = vbGreen
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 3 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbRed
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 4 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbYellow
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 5 Then
Picture3.BackColor = vbBlack
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
Dim cur1 As Integer
For cur1 = 0 To Me.mnuColor.UBound
Me.mnuColor(cur1).Checked = False
Next
Me.mnuColor(Combo2.ListIndex).Checked = True
For cur1 = 0 To Me.mnuBGC.UBound
mnuBGC(cur1).Checked = False
Next
Me.mnuBGC(Combo2.ListIndex).Checked = True
For cur1 = 0 To Me.mnuSize.UBound
Me.mnuSize(cur1).Checked = False
Next
Me.mnuSize(Combo1.ListIndex).Checked = True
For cur1 = 0 To Me.mnuSizeView.UBound
mnuSizeView(cur1).Checked = False
Next
Me.mnuSizeView(Combo1.ListIndex).Checked = True
Select Case Me.mnuSaveExit.Checked
Case True
SaveSetting "Icon Tools", "Options", "DrawSize", Me.Combo1.ListIndex
SaveSetting "Icon Tools", "Options", "BackColor", Trim(Str(Me.Combo2.ListIndex))
Case False
SaveSetting "Icon Tools", "Options", "DrawSize", 0
SaveSetting "Icon Tools", "Options", "BackColor", 0
End Select
End Sub
Private Sub Combo2_Click()
On Error Resume Next
With Me.StateBar1
.Alignment = 0
.Caption = "页面设置更改"
End With
If Combo2.ListIndex = 0 Then
Picture3.BackColor = vbWhite
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 1 Then
Picture3.BackColor = vbBlue
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 2 Then
Picture3.BackColor = vbGreen
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 3 Then
Picture3.BackColor = vbRed
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 4 Then
Picture3.BackColor = vbYellow
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 5 Then
Picture3.BackColor = vbBlack
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
Dim cur1 As Integer
For cur1 = 0 To Me.mnuColor.UBound
Me.mnuColor(cur1).Checked = False
Next
Me.mnuColor(Combo2.ListIndex).Checked = True
For cur1 = 0 To Me.mnuBGC.UBound
mnuBGC(cur1).Checked = False
Next
Me.mnuBGC(Combo2.ListIndex).Checked = True
For cur1 = 0 To Me.mnuSize.UBound
Me.mnuSize(cur1).Checked = False
Next
Me.mnuSize(Combo1.ListIndex).Checked = True
For cur1 = 0 To Me.mnuSizeView.UBound
mnuSizeView(cur1).Checked = False
Next
Me.mnuSizeView(Combo1.ListIndex).Checked = True
Select Case Me.mnuSaveExit.Checked
Case True
SaveSetting "Icon Tools", "Options", "DrawSize", Me.Combo1.ListIndex
SaveSetting "Icon Tools", "Options", "BackColor", Trim(Str(Me.Combo2.ListIndex))
Case False
SaveSetting "Icon Tools", "Options", "DrawSize", 0
SaveSetting "Icon Tools", "Options", "BackColor", 0
End Select
End Sub
Private Sub Command1_Click()
On Error Resume Next
frmOpenFile.Show 1
End Sub
Private Sub Command2_Click()
On Error Resume Next
frmSaveFile.Show 1
End Sub
Private Sub Command5_Click()
On Error Resume Next
Dim ans As Integer
ans = vbYes
If ans = vbYes Then
Select Case Me.mnuSaveExit.Checked
Case True
SaveSetting "Icon Tools", "Options", "DrawSize", Me.Combo1.ListIndex
SaveSetting "Icon Tools", "Options", "BackColor", Me.Combo2.ListIndex
Case False
SaveSetting "Icon Tools", "Options", "DrawSize", 0
SaveSetting "Icon Tools", "Options", "BackColor", 0
End Select
Unload Me
Unload Form2
Unload frmAbout
Unload frmOpenFile
Unload FrmSaveConverted
Unload frmSaveFile
End
Else
Exit Sub
End If
End Sub
Private Sub Command6_Click()
On Error Resume Next
cur = cur - 1
If cur > 0 Then
Command6.Enabled = True
With Me.Label5
.Enabled = True
.Caption = cur
.Alignment = 2
End With
With Me.StateBar1
.Alignment = 0
.Caption = "指针上移"
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
Me.Command2.Enabled = True
Me.Command6.Enabled = True
Me.Command7.Enabled = True
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
If Combo1.ListIndex = 0 Then
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
DrawIconEx Picture3.hdc, 0, 0, hIco, 0, 0, 0, 0, DI_NORMAL Or DI_DEFAULTSIZE
End If
With Me.Picture3
.Refresh
End With
ElseIf cur = 0 Then
Command6.Enabled = False
With Me.StateBar1
.Alignment = 0
.Caption = "指针归零"
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
Me.Command2.Enabled = True
Me.Command7.Enabled = True
Command6.Enabled = False
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
If Combo1.ListIndex = 0 Then
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
DrawIconEx Picture3.hdc, 0, 0, hIco, 0, 0, 0, 0, DI_NORMAL Or DI_DEFAULTSIZE
End If
With Me.Picture3
.Refresh
End With
With Me.Label5
.Enabled = True
.Caption = cur
.Alignment = 2
End With
End If
End Sub
Private Sub Command7_Click()
On Error Resume Next
cur = cur + 1
If cur < Label4.Caption Then
Command7.Enabled = True
With Me.StateBar1
.Alignment = 0
.Caption = "指针下移"
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
Me.Command2.Enabled = True
Me.Command6.Enabled = True
Me.Command7.Enabled = True
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
If Combo1.ListIndex = 0 Then
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
DrawIconEx Picture3.hdc, 0, 0, hIco, 0, 0, 0, 0, DI_NORMAL Or DI_DEFAULTSIZE
End If
With Me.Picture3
.Refresh
End With
With Me.Label5
.Enabled = True
.Caption = cur
.Alignment = 2
End With
Else
With Me.StateBar1
.Alignment = 0
.Caption = "指针到底"
End With
Command7.Enabled = False
End If
If cur = Val(Label4.Caption) - 1 Then
Command7.Enabled = False
With Me.StateBar1
.Alignment = 0
.Caption = "指针到底"
End With
End If
End Sub
Private Sub Form_Activate()
On Error Resume Next
With Me.StateBar1
.Alignment = 0
.Caption = "准备就绪"
End With
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyO And Shift = vbCtrlMask Then
On Error Resume Next
frmOpenFile.Show 1
End If
End Sub
Private Sub Form_Load()
On Error Resume Next
Dim a As Integer
a = Val(GetSetting("Icon Tools", "Options", "BackColor", 0))
Me.mnuCopy2.Enabled = False
With Me.mnuSaveExit
.Checked = True
End With
With Me.StateBar1
.Alignment = 0
.Caption = "准备就绪"
End With
With Me.Label4
.Alignment = 2
.Caption = ""
End With
With Me.Image1
.Visible = False
.Enabled = False
End With
With Me.Label5
.Enabled = False
.Caption = ""
.Alignment = 2
End With
Me.mnuCopy.Enabled = False
Me.mnuJump.Enabled = False
With Me
.mnuCopy.Enabled = False
.Left = Screen.Width / 2 - .Width / 2
.TOp = Screen.Height / 2 - .Height / 2
.Height = 4590
.Width = 6135
.Icon = Me.Image1.Picture
.Picture = LoadPicture()
End With
Me.mnuClose.Enabled = False
Me.Command2.Enabled = False
Me.Command6.Enabled = False
Me.Command7.Enabled = False
Me.mnuCopy2.Enabled = False
Me.mnuJump.Enabled = False
With Picture1
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
Me.Text1.Locked = True
Me.Label4.Caption = ""
With Me
.KeyPreview = True
.Command5.Cancel = True
.Command1.Default = True
End With
Me.Command2.Enabled = False
Me.mnuSave.Enabled = False
Me.Command6.Enabled = False
Me.Command7.Enabled = False
Label4.Caption = ""
With Me.Combo1
.Enabled = True
End With
With Me.Combo2
.Enabled = True
End With
With Me
.Height = 4845
.Width = 6090
End With
With Me.Combo1
.ListIndex = Val(GetSetting("Icon Tools", "Options", "DrawSize", 0))
.Enabled = True
End With
With Me.Combo2
.ListIndex = a
If .ListIndex = -1 Then
.ListIndex = 0
End If
.Enabled = True
End With
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Dim ans As Integer
ans = MsgBox("确定退出程序?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
Select Case Me.mnuSaveExit.Checked
Case True
SaveSetting "Icon Tools", "Options", "DrawSize", Me.Combo1.ListIndex
SaveSetting "Icon Tools", "Options", "BackColor", Me.Combo2.ListIndex
Case False
SaveSetting "Icon Tools", "Options", "DrawSize", 0
SaveSetting "Icon Tools", "Options", "BackColor", 0
End Select
Unload Me
Unload Form2
Unload frmAbout
Unload frmOpenFile
Unload FrmSaveConverted
Unload frmSaveFile
Else
Cancel = 666 + 666
Exit Sub
End If
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Select Case Me.mnuSaveExit.Checked
Case True
SaveSetting "Icon Tools", "Options", "DrawSize", Me.Combo1.ListIndex
SaveSetting "Icon Tools", "Options", "BackColor", Me.Combo2.ListIndex
Case False
SaveSetting "Icon Tools", "Options", "DrawSize", 0
SaveSetting "Icon Tools", "Options", "BackColor", 0
End Select
Unload Me
Unload Form2
Unload frmAbout
Unload frmOpenFile
Unload FrmSaveConverted
Unload frmSaveFile
End Sub
Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = 2 Then
PopupMenu Me.mnuFile
Else
Exit Sub
End If
End Sub
Private Sub Label5_Click()
On Error Resume Next
Dim strCur As String
Dim lngCur As Long
strCur = InputBox$("请输入您想要跳转到的位置" & vbCrLf & "有效范围:" & vbCrLf & "0~" & Label4.Caption - 1, "Jump")
If Trim(strCur) = "" Then
Exit Sub
End If
lngCur = Val(strCur)
If lngCur >= 0 And lngCur <= Val(Label4.Caption) - 1 Then
cur = lngCur
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), lngCur)
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
DrawIconEx Picture3.hdc, 0, 0, hIco, 0, 0, 0, 0, DI_NORMAL Or DI_DEFAULTSIZE
With Me.Picture3
.Refresh
End With
With Me.Label5
.Enabled = True
.Caption = lngCur
.Alignment = 2
End With
If lngCur = 0 Then
With Me.Command6
.Enabled = False
End With
With Me.Command7
.Enabled = True
End With
End If
If lngCur = Val(Label4.Caption) - 1 Then
With Me.Command7
.Enabled = False
End With
With Me.Command6
.Enabled = True
End With
With Me.StateBar1
.Alignment = 0
.Caption = "指针跳转"
End With
End If
cur = lngCur
ElseIf lngCur < 0 Or lngCur > Label4.Caption - 1 Then
With Me.Label5
.Enabled = True
.Alignment = 2
.Caption = cur
End With
MsgBox "请输入有效的图标指针序号", vbExclamation, "Error"
With Me.Label5
.Enabled = True
.Alignment = 2
.Caption = cur
End With
End If
End Sub
Private Sub mnuAbout_Click()
On Error Resume Next
frmAbout.Show 1
End Sub
Private Sub mnuBGC_Click(Index As Integer)
On Error Resume Next
With Me.StateBar1
.Alignment = 0
.Caption = "页面设置更改"
End With
If Index = 0 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbWhite
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Index = 1 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbBlue
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Index = 2 Then
Picture3.BackColor = vbGreen
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Index = 3 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbRed
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Index = 4 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbYellow
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Index = 5 Then
Picture3.BackColor = vbBlack
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
Combo2.ListIndex = Index
Dim cur1 As Integer
For cur1 = 0 To Me.mnuColor.UBound
Me.mnuColor(cur1).Checked = False
Next
Me.mnuColor(Index).Checked = True
For cur1 = 0 To Me.mnuBGC.UBound
mnuBGC(cur1).Checked = False
Next
Me.mnuBGC(Index).Checked = True
End Sub
Private Sub mnuClear_Click()
On Error GoTo ep
Dim ans As Integer
ans = MsgBox("确认清除(复位)保存的设置吗?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
SaveSetting "Icon Tools", "Options", "DrawSize", 0
SaveSetting "Icon Tools", "Options", "BackColor", 0
MsgBox "设置复位成功!", vbExclamation, "Info"
Else
Exit Sub
End If
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
SaveSetting "Icon Tools", "Options", "DrawSize", 0
SaveSetting "Icon Tools", "Options", "BackColor", 0
Exit Sub
End Sub
Private Sub mnuClose_Click()
On Error Resume Next
Dim ans As Integer
ans = MsgBox("确定关闭当前打开的Win32PE文件?", vbQuestion + vbYesNo, "Ask")
If ans = vbYes Then
cod = True
With Me
.Tag = ""
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
With Me.Text1
.Text = ""
.Locked = True
End With
Me.Command2.Enabled = False
Me.mnuSave.Enabled = False
Me.Command6.Enabled = False
Me.Command7.Enabled = False
Me.mnuClose.Enabled = False
Me.mnuJump.Enabled = False
Me.mnuCopy2.Enabled = False
With Me.StateBar1
.Alignment = 0
.Caption = "文件关闭"
End With
Me.mnuCopy.Enabled = False
Me.Label4.Caption = ""
With Me.Label5
.Enabled = False
.Caption = ""
.Alignment = 2
End With
cod = False
Else
cod = False
Exit Sub
With Me.Combo1
.Enabled = True
.ListIndex = 1
End With
End If
End Sub
Private Sub mnuColor_Click(Index As Integer)
On Error Resume Next
With Me.StateBar1
.Alignment = 0
.Caption = "页面设置更改"
End With
If Index = 0 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbWhite
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Index = 1 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbBlue
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Index = 2 Then
Picture3.BackColor = vbGreen
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Index = 3 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbRed
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Index = 4 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbYellow
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Index = 5 Then
Picture3.BackColor = vbBlack
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Combo1.ListIndex = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo1.ListIndex = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
Combo2.ListIndex = Index
Dim cur1 As Integer
For cur1 = 0 To Me.mnuColor.UBound
Me.mnuColor(cur1).Checked = False
Next
Me.mnuColor(Index).Checked = True
For cur1 = 0 To Me.mnuBGC.UBound
mnuBGC(cur1).Checked = False
Next
Me.mnuBGC(Index).Checked = True
Select Case Me.mnuSaveExit.Checked
Case True
SaveSetting "Icon Tools", "Options", "DrawSize", Me.Combo1.ListIndex
SaveSetting "Icon Tools", "Options", "BackColor", Me.Combo2.ListIndex
Case False
SaveSetting "Icon Tools", "Options", "DrawSize", 0
SaveSetting "Icon Tools", "Options", "BackColor", 0
End Select
End Sub
Private Sub mnuConvert_Click()
On Error Resume Next
frmAlphaIconCreator.Show 1
End Sub
Private Sub mnuCopy_Click()
On Error GoTo ep
With Clipboard
.Clear
.SetData Me.Picture3.Image
End With
MsgBox "已经将图标发送到剪切板,可以在任何绘图软件中使用CTRL+V或其它粘贴快捷键粘贴!", vbExclamation, "Info"
With Me.StateBar1
.Alignment = 0
.Caption = "图标复制成功..."
End With
Exit Sub
Exit Sub
ep:
MsgBox "复制图片失败,错误" & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
With Me.StateBar1
.Alignment = 0
.Caption = "图标复制失败..."
End With
Exit Sub
Exit Sub
End Sub
Private Sub mnuCopy2_Click()
On Error GoTo ep
With Clipboard
.Clear
.SetData Me.Picture3.Image
End With
MsgBox "已经将图标发送到剪切板,可以在任何绘图软件中使用CTRL+V或其它粘贴快捷键粘贴!", vbExclamation, "Info"
With Me.StateBar1
.Alignment = 0
.Caption = "图标复制成功..."
End With
Exit Sub
ep:
MsgBox "复制图片失败,错误" & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
With Me.StateBar1
.Alignment = 0
.Caption = "图标复制失败..."
End With
Exit Sub
End Sub
Private Sub mnuCVS_Click()
On Error Resume Next
Exit Sub
Form2.Show 1
End Sub
Private Sub mnuExit_Click()
On Error Resume Next
Dim ans As Integer
ans = vbYes
If ans = vbYes Then
Select Case Me.mnuSaveExit.Checked
Case Is = True
SaveSetting "Icon Tools", "Options", "DrawSize", Me.Combo1.ListIndex
SaveSetting "Icon Tools", "Options", "BackColor", Me.Combo2.ListIndex
Case Is = False
SaveSetting "Icon Tools", "Options", "DrawSize", 0
SaveSetting "Icon Tools", "Options", "BackColor", 0
End Select
Unload Me
Unload Form2
Unload frmAbout
Unload frmOpenFile
Unload FrmSaveConverted
Unload frmSaveFile
Else
Exit Sub
End If
End Sub
Private Sub mnuHide_Click()
On Error Resume Next
MsgBox "执行时发生错误,过程终止" & vbCrLf & "错误245:Windows菜单管理器无法处理函数", vbCritical, "Error"
Exit Sub
Dim ans As Integer
ans = MsgBox("确实要把它从菜单栏隐藏吗?" & vbCrLf & "隐藏后,您可以从'文件'菜单找到它.", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Me.mnuCVS.Visible = False
Else
Me.mnuCVS.Visible = True
End If
End Sub
Private Sub mnuJump_Click()
On Error Resume Next
Dim strCur As String
Dim lngCur As Long
strCur = InputBox$("请输入您想要跳转到的位置" & vbCrLf & "有效范围:" & vbCrLf & "0~" & Label4.Caption - 1, "Jump")
If Trim(strCur) = "" Then
Exit Sub
End If
lngCur = Val(strCur)
If lngCur >= 0 And lngCur <= Val(Label4.Caption) - 1 Then
cur = lngCur
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), lngCur)
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
DrawIconEx Picture3.hdc, 0, 0, hIco, 0, 0, 0, 0, DI_NORMAL Or DI_DEFAULTSIZE
With Me.Picture3
.Refresh
End With
With Me.Label5
.Enabled = True
.Caption = lngCur
.Alignment = 2
End With
If lngCur = 0 Then
With Me.Command6
.Enabled = False
End With
With Me.Command7
.Enabled = True
End With
End If
If lngCur = Val(Label4.Caption) - 1 Then
With Me.Command7
.Enabled = False
End With
With Me.Command6
.Enabled = True
End With
End If
cur = lngCur
With Me.StateBar1
.Alignment = 0
.Caption = "指针跳转"
End With
ElseIf lngCur < 0 Or lngCur > Label4.Caption - 1 Then
With Me.Label5
.Enabled = True
.Alignment = 2
.Caption = cur
End With
MsgBox "请输入有效的图标指针序号", vbExclamation, "Error"
With Me.Label5
.Enabled = True
.Alignment = 2
.Caption = cur
End With
End If
End Sub
Private Sub mnuOpen_Click()
On Error Resume Next
frmOpenFile.Show 1
End Sub
Private Sub mnuRun_Click()
On Error Resume Next
frmAlphaIconCreator.Show 1
End Sub
Private Sub mnuSave_Click()
On Error Resume Next
frmSaveFile.Show 1
End Sub
Private Sub mnuSaveExit_Click()
On Error Resume Next
If mnuSaveExit.Checked = True Then
mnuSaveExit.Checked = False
SaveSetting "Icon Tools", "Options", "DrawSize", 0
SaveSetting "Icon Tools", "Options", "BackColor", 0
Exit Sub
End If
If mnuSaveExit.Checked = False Then
mnuSaveExit.Checked = True
SaveSetting "Icon Tools", "Options", "DrawSize", Me.Combo1.ListIndex
SaveSetting "Icon Tools", "Options", "BackColor", Me.Combo2.ListIndex
Exit Sub
End If
End Sub
Private Sub mnuShowCurrent_Click()
On Error Resume Next
Dim DS As Integer
Dim BC As Integer
DS = Val(GetSetting("Icon Tools", "Options", "DrawSize", 0))
BC = Val(GetSetting("Icon Tools", "Options", "BackColor", 0))
If Me.mnuSaveExit.Checked = False Then
MsgBox "当前不允许保存设置,下面是默认选项:" & vbCrLf & "                                                                                   " & vbCrLf & "绘图大小(像素):(默认)使用系统指定的值" & vbCrLf & "绘图背景色:白色", vbInformation, "Info"
Exit Sub
End If
If Me.mnuSaveExit.Checked = True Then
MsgBox "当前允许保存设置,下面是当前选项:" & vbCrLf & "                                                                                   " & vbCrLf & "绘图大小(像素):" & Combo1.List(DS) & vbCrLf & "绘图背景色:" & Combo2.List(BC), vbInformation, "Info"
Exit Sub
End If
End Sub
Private Sub mnuSize_Click(Index As Integer)
On Error Resume Next
With Me.StateBar1
.Alignment = 0
.Caption = "页面设置更改"
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo2.ListIndex = 0 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbWhite
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 1 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbBlue
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 2 Then
Picture3.BackColor = vbGreen
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 3 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbRed
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 4 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbYellow
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 5 Then
Picture3.BackColor = vbBlack
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 0 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbWhite
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 1 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbBlue
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 2 Then
Picture3.BackColor = vbGreen
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 3 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbRed
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 4 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbYellow
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 5 Then
Picture3.BackColor = vbBlack
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
Combo1.ListIndex = Index
Dim cur1 As Integer
For cur1 = 0 To Me.mnuSize.UBound
Me.mnuSize(cur1).Checked = False
Next
Me.mnuSize(Index).Checked = True
For cur1 = 0 To Me.mnuSizeView.UBound
mnuSizeView(cur1).Checked = False
Next
Me.mnuSizeView(Index).Checked = True
Select Case Me.mnuSaveExit.Checked
Case True
SaveSetting "Icon Tools", "Options", "DrawSize", Me.Combo1.ListIndex
SaveSetting "Icon Tools", "Options", "BackColor", Me.Combo2.ListIndex
Case False
SaveSetting "Icon Tools", "Options", "DrawSize", 0
SaveSetting "Icon Tools", "Options", "BackColor", 0
End Select
End Sub
Private Sub mnuSizeView_Click(Index As Integer)
On Error Resume Next
With Me.StateBar1
.Alignment = 0
.Caption = "页面设置更改"
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Combo2.ListIndex = 0 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbWhite
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 1 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbBlue
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 2 Then
Picture3.BackColor = vbGreen
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 3 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbRed
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 4 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbYellow
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 5 Then
Picture3.BackColor = vbBlack
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 0 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbWhite
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 1 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbBlue
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 2 Then
Picture3.BackColor = vbGreen
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 3 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbRed
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 4 Then
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
Picture3.BackColor = vbYellow
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
If Combo2.ListIndex = 5 Then
Picture3.BackColor = vbBlack
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
If Index = 0 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 1 Then
With Picture3
.Height = 16
.Width = 16
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 2 Then
With Picture3
.Height = 32
.Width = 32
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 3 Then
With Picture3
.Height = 48
.Width = 48
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 4 Then
With Picture3
.Height = 64
.Width = 64
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
If Index = 5 Then
With Picture3
.Height = 128
.Width = 128
.AutoRedraw = True
End With
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), cur)
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
With Me.Picture3
.Refresh
End With
End If
End If
Combo1.ListIndex = Index
Dim cur1 As Integer
For cur1 = 0 To Me.mnuSize.UBound
Me.mnuSize(cur1).Checked = False
Next
Me.mnuSize(Index).Checked = True
For cur1 = 0 To Me.mnuSizeView.UBound
mnuSizeView(cur1).Checked = False
Next
Me.mnuSizeView(Index).Checked = True
Select Case Me.mnuSaveExit.Checked
Case True
SaveSetting "Icon Tools", "Options", "DrawSize", Me.Combo1.ListIndex
SaveSetting "Icon Tools", "Options", "BackColor", Me.Combo2.ListIndex
Case False
SaveSetting "Icon Tools", "Options", "DrawSize", 0
SaveSetting "Icon Tools", "Options", "BackColor", 0
End Select
End Sub
Private Sub Picture3_Click()
On Error GoTo ep
If Me.Command2.Enabled = True Then
With Clipboard
.Clear
.SetData Me.Picture3.Image
End With
MsgBox "已经将图标发送到剪切板,可以在任何绘图软件中使用CTRL+V或其它粘贴快捷键粘贴!", vbExclamation, "Info"
With Me.StateBar1
.Alignment = 0
.Caption = "图标复制成功..."
End With
Exit Sub
Exit Sub
Else
Exit Sub
End If
Exit Sub
ep:
MsgBox "复制图片失败,错误" & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
With Me.StateBar1
.Alignment = 0
.Caption = "图标复制失败..."
End With
Exit Sub
Exit Sub
End Sub
Private Sub Picture3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If Button = 2 Then
PopupMenu Me.mnuFile
Else
Exit Sub
End If
End Sub
Private Sub Text1_Change()
On Error Resume Next
If cod = False Then
tmp = ExtractIcon(App.hInstance, Trim(Me.Tag), -1)
If tmp > 0 Then
Label4.Caption = tmp
Else
MsgBox "提取图标失败!", vbCritical, "Error"
Me.Command2.Enabled = False
Me.Command6.Enabled = False
Me.Command7.Enabled = False
Me.mnuSave.Enabled = False
Me.mnuClose.Enabled = False
Me.mnuCopy.Enabled = False
Me.mnuCopy2.Enabled = False
Me.mnuJump.Enabled = False
With Me.StateBar1
.Alignment = 0
.Caption = "图标提取失败"
End With
Label4.Caption = ""
With Me.Label5
.Enabled = False
.Caption = ""
.Alignment = 2
End With
Exit Sub
End If
hIco = ExtractIcon(App.hInstance, Trim(Me.Tag), 0)
If hIco <> 0 Then
Me.Command2.Enabled = True
Me.Command6.Enabled = False
Me.Command7.Enabled = True
Me.mnuClose.Enabled = True
Me.mnuSave.Enabled = True
Me.mnuCopy2.Enabled = True
Me.mnuCopy.Enabled = True
Me.mnuJump.Enabled = False
Me.mnuJump.Enabled = True
With Me.StateBar1
.Alignment = 0
.Caption = "Win32PE文件打开成功"
End With
If Val(Label4.Caption) = 1 Then
Me.Command7.Enabled = False
Me.mnuJump.Enabled = False
With Me.Label5
.Alignment = 2
.Enabled = False
.Caption = "0"
End With
End If
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
DrawIconEx Picture3.hdc, 0, 0, hIco, Picture3.ScaleWidth, Picture3.ScaleHeight, 0, 0, DI_NORMAL
If Combo1.ListIndex = 0 Then
With Me.Picture3
.Cls
.Picture = LoadPicture()
.AutoRedraw = True
End With
DrawIconEx Picture3.hdc, 0, 0, hIco, 0, 0, 0, 0, DI_NORMAL Or DI_DEFAULTSIZE
End If
With Me.Picture3
.Refresh
End With
cur = 0
With Me.Label5
.Enabled = True
.Caption = cur
.Alignment = 2
End With
If Val(Label4.Caption) = 1 Then
With Me.Label5
.Enabled = False
.Caption = cur
.Alignment = 2
End With
End If
Else
MsgBox "图标提取失败!", vbCritical, "Error"
Me.Command2.Enabled = False
Me.Command6.Enabled = False
Me.Command7.Enabled = False
With Me.StateBar1
.Alignment = 0
.Caption = "图标提取失败"
End With
Me.mnuSave.Enabled = False
Me.mnuCopy.Enabled = False
Me.mnuClose.Enabled = False
Me.mnuJump.Enabled = False
Me.mnuCopy2.Enabled = False
Label4.Caption = ""
With Me.Label5
.Enabled = False
.Caption = ""
.Alignment = 2
End With
End If
End If
End Sub
