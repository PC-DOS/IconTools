VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Begin VB.Form frmSaveFile 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保存图标"
   ClientHeight    =   6225
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13395
   ControlBox      =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6225
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture2 
      Height          =   6360
      Left            =   7845
      ScaleHeight     =   6300
      ScaleWidth      =   11190
      TabIndex        =   14
      Top             =   6570
      Visible         =   0   'False
      Width           =   11250
      Begin VB.PictureBox picTab 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4635
         Index           =   2
         Left            =   4530
         ScaleHeight     =   4635
         ScaleWidth      =   5415
         TabIndex        =   40
         Top             =   750
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CommandButton cmdStartOver 
            BackColor       =   &H80000005&
            Caption         =   "重執行"
            Height          =   375
            Left            =   120
            TabIndex        =   41
            Top             =   840
            Width           =   1275
         End
         Begin VB.Label lblInfo 
            BackStyle       =   0  'Transparent
            Caption         =   "保存圖標成功"
            Height          =   675
            Left            =   120
            TabIndex        =   42
            Top             =   180
            Width           =   5115
         End
      End
      Begin VB.PictureBox picTab 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4635
         Index           =   1
         Left            =   3900
         ScaleHeight     =   4635
         ScaleWidth      =   5415
         TabIndex        =   26
         Top             =   735
         Visible         =   0   'False
         Width           =   5415
         Begin VB.CheckBox chkSize 
            BackColor       =   &H80000005&
            Caption         =   "16 x 16"
            Height          =   195
            Index           =   0
            Left            =   1260
            TabIndex        =   37
            Top             =   120
            Value           =   1  'Checked
            Width           =   3915
         End
         Begin VB.CheckBox chkSize 
            BackColor       =   &H80000005&
            Caption         =   "32 x 32"
            Height          =   195
            Index           =   2
            Left            =   1260
            TabIndex        =   36
            Top             =   594
            Value           =   1  'Checked
            Width           =   3915
         End
         Begin VB.CheckBox chkSize 
            BackColor       =   &H80000005&
            Caption         =   "48 x 48"
            Height          =   195
            Index           =   3
            Left            =   1260
            TabIndex        =   35
            Top             =   831
            Value           =   1  'Checked
            Width           =   3915
         End
         Begin VB.TextBox txtFileName 
            Height          =   285
            Left            =   1260
            TabIndex        =   34
            Top             =   3180
            Width           =   3975
         End
         Begin VB.CommandButton cmdPickOutput 
            BackColor       =   &H80000005&
            Caption         =   "选择..."
            Height          =   375
            Left            =   1260
            TabIndex        =   33
            Top             =   3540
            Width           =   1095
         End
         Begin VB.CheckBox chkSize 
            BackColor       =   &H80000005&
            Caption         =   "24 x 24"
            Height          =   195
            Index           =   1
            Left            =   1260
            TabIndex        =   32
            Top             =   357
            Width           =   3915
         End
         Begin VB.CheckBox chkSize 
            BackColor       =   &H80000005&
            Caption         =   "64 x 64"
            Height          =   195
            Index           =   4
            Left            =   1260
            TabIndex        =   31
            Top             =   1068
            Width           =   3915
         End
         Begin VB.ListBox lstCustom 
            Height          =   1140
            Left            =   6120
            TabIndex        =   30
            Top             =   1380
            Visible         =   0   'False
            Width           =   3735
         End
         Begin VB.CommandButton cmdAddCustom 
            BackColor       =   &H80000005&
            Caption         =   "添加..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   7650
            TabIndex        =   29
            Top             =   2640
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CommandButton cmdRemoveCustom 
            BackColor       =   &H80000005&
            Caption         =   "刪除..."
            Enabled         =   0   'False
            Height          =   375
            Left            =   8760
            TabIndex        =   28
            Top             =   2640
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox chkSize 
            BackColor       =   &H80000005&
            Caption         =   "128 x 128"
            Height          =   195
            Index           =   5
            Left            =   1260
            TabIndex        =   27
            Top             =   1305
            Width           =   3915
         End
         Begin VB.Label lblSizes 
            BackStyle       =   0  'Transparent
            Caption         =   "大小"
            Height          =   195
            Left            =   0
            TabIndex        =   39
            Top             =   120
            Width           =   1215
         End
         Begin VB.Label lblOutputFile 
            BackStyle       =   0  'Transparent
            Caption         =   "文件名"
            Height          =   195
            Left            =   360
            TabIndex        =   38
            Top             =   3240
            Width           =   795
         End
      End
      Begin VB.CommandButton cmdCancel 
         BackColor       =   &H80000016&
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   375
         Left            =   7995
         TabIndex        =   25
         Top             =   5670
         Width           =   1275
      End
      Begin VB.CommandButton cmdNext 
         BackColor       =   &H80000016&
         Caption         =   "下一步(&N)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6615
         TabIndex        =   24
         Top             =   5670
         Width           =   1275
      End
      Begin VB.CommandButton cmdBack 
         BackColor       =   &H80000016&
         Caption         =   "後退(&B)"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5295
         TabIndex        =   23
         Top             =   5670
         Width           =   1275
      End
      Begin VB.PictureBox picTab 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   4575
         Index           =   0
         Left            =   3900
         ScaleHeight     =   4575
         ScaleWidth      =   5415
         TabIndex        =   15
         Top             =   735
         Width           =   5415
         Begin VB.PictureBox picSource 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00F0F0F0&
            ForeColor       =   &H80000008&
            Height          =   2295
            Left            =   1320
            ScaleHeight     =   2265
            ScaleWidth      =   2265
            TabIndex        =   20
            Top             =   780
            Width           =   2295
         End
         Begin VB.TextBox txtSourceImage 
            Height          =   285
            Left            =   1320
            TabIndex        =   19
            Top             =   60
            Width           =   3975
         End
         Begin VB.CommandButton cmdPickInput 
            BackColor       =   &H80000005&
            Caption         =   "瀏覽..."
            Height          =   375
            Left            =   1320
            TabIndex        =   18
            Top             =   360
            Width           =   1095
         End
         Begin VB.PictureBox picTransparentColour 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1320
            ScaleHeight     =   285
            ScaleWidth      =   3945
            TabIndex        =   17
            Top             =   3360
            Width           =   3975
         End
         Begin VB.CommandButton cmdPickColour 
            BackColor       =   &H80000005&
            Caption         =   "選擇..."
            Height          =   375
            Left            =   1320
            TabIndex        =   16
            Top             =   3720
            Width           =   1095
         End
         Begin VB.Label lblSourceImage 
            BackStyle       =   0  'Transparent
            Caption         =   "源圖片"
            Height          =   195
            Left            =   60
            TabIndex        =   22
            Top             =   120
            Width           =   1155
         End
         Begin VB.Label lblTransparentColour 
            BackStyle       =   0  'Transparent
            Caption         =   "透明色"
            Height          =   435
            Left            =   60
            TabIndex        =   21
            Top             =   3360
            Width           =   1155
         End
      End
      Begin VB.Image Image1 
         Height          =   5400
         Left            =   90
         Picture         =   "frmSaveFile.frx":0000
         Stretch         =   -1  'True
         Top             =   0
         Width           =   3795
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000016&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000016&
         Height          =   2145
         Left            =   0
         Top             =   5400
         Width           =   10845
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "圖標創建向導"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   24.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   5
         Left            =   3930
         TabIndex        =   43
         Top             =   105
         Width           =   5295
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5010
      Left            =   6975
      ScaleHeight     =   5010
      ScaleWidth      =   6375
      TabIndex        =   12
      Top             =   30
      Width           =   6375
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         CausesValidation=   0   'False
         Height          =   4965
         Left            =   0
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   0
         Width           =   6360
         ExtentX         =   11218
         ExtentY         =   8758
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   0
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
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
   End
   Begin VB.FileListBox File1 
      Height          =   4770
      Left            =   2895
      TabIndex        =   11
      Top             =   255
      Width           =   4080
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frmSaveFile.frx":B98CE
      Left            =   945
      List            =   "frmSaveFile.frx":B98DB
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   5415
      Width           =   12405
   End
   Begin VB.DriveListBox Drive1 
      Height          =   300
      Left            =   720
      TabIndex        =   4
      Top             =   15
      Width           =   2130
   End
   Begin VB.DirListBox Dir1 
      Height          =   4290
      Left            =   15
      TabIndex        =   3
      Top             =   720
      Width           =   2865
   End
   Begin VB.TextBox Text1 
      Height          =   300
      Left            =   945
      TabIndex        =   2
      Top             =   5040
      Width           =   12405
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存(&S)"
      Default         =   -1  'True
      Height          =   420
      Left            =   11025
      TabIndex        =   1
      Top             =   5760
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   12210
      TabIndex        =   0
      Top             =   5760
      Width           =   1125
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件类型"
      Height          =   180
      Index           =   4
      Left            =   60
      TabIndex        =   9
      Top             =   5460
      Width           =   720
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "驱动器"
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   8
      Top             =   60
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件夹"
      Height          =   180
      Index           =   1
      Left            =   30
      TabIndex        =   7
      Top             =   420
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件"
      Height          =   180
      Index           =   2
      Left            =   2895
      TabIndex        =   6
      Top             =   45
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "文件名"
      Height          =   180
      Index           =   3
      Left            =   45
      TabIndex        =   5
      Top             =   5100
      Width           =   540
   End
End
Attribute VB_Name = "frmSaveFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cFI As New cFileIcon
Dim qAns As Integer
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)

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
Dim bth As String
Private Sub Command1_Click()

On Error GoTo ep
Dim ans As Integer
Me.Dir1.Enabled = False
Me.File1.Enabled = False
Me.Drive1.Enabled = False
Me.Command1.Enabled = False
Me.Command2.Enabled = False
Me.Combo1.Enabled = False
If Trim(Text1.Text) = "" Then
MsgBox "请输入文件名!", vbCritical, "Error"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
Exit Sub
End If
Select Case Right(Dir1.Path, 1)
Case "\"
bth = Dir1.Path
Case Else
bth = Dir1.Path & "\"
End Select
If Combo1.ListIndex = 0 Then
If Trim(Text1.Text) <> "" Then
If Dir(bth & Text1.Text & ".ico") = "" Then
Select Case Combo1.ListIndex
Case 0
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".suicune"
      Set m_pic = Nothing
      cmdNext.Enabled = False
      'Suicune
      qAns = MsgBox("是否保留圖標背景色?", vbQuestion + vbYesNo, "Ask")
      If qAns = vbNo Then
      txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      Select Case Form1.Combo2.ListIndex
      '白色
      '蓝色
      '绿色
      '红色
      '黄色
      '黑色
      Case 0
      setTransparentColour vbWhite
      Case 1
      setTransparentColour vbBlue
      Case 2
      setTransparentColour vbGreen
      Case 3
      setTransparentColour vbRed
      Case 4
      setTransparentColour vbYellow
      Case 5
      setTransparentColour vbBlack
      End Select
      'setTransparentColour picTransparentColour.BackColor
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      Else
            txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      'setTransparentColour RGB(0, 245, 245)
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      End If
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
'MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
ans = MsgBox("文件已经存在,是否替换?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill bth & Text1.Text & ".ico"
Select Case Combo1.ListIndex
Case 0
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".suicune"
      Set m_pic = Nothing
      cmdNext.Enabled = False
      txtSourceImage.Text = ""
      'Suicune
      qAns = MsgBox("是否保留圖標背景色?", vbQuestion + vbYesNo, "Ask")
      If qAns = vbNo Then
      txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      Select Case Form1.Combo2.ListIndex
      '白色
      '蓝色
      '绿色
      '红色
      '黄色
      '黑色
      Case 0
      setTransparentColour vbWhite
      Case 1
      setTransparentColour vbBlue
      Case 2
      setTransparentColour vbGreen
      Case 3
      setTransparentColour vbRed
      Case 4
      setTransparentColour vbYellow
      Case 5
      setTransparentColour vbBlack
      End Select
      'setTransparentColour picTransparentColour.BackColor
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      Else
            txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      'setTransparentColour RGB(0, 245, 245)
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      End If
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
'MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
Exit Sub
End If
End If
End If
End If
If Combo1.ListIndex = 1 Then
If Trim(Text1.Text) <> "" Then
If Dir(bth & Text1.Text & ".jpg") = "" Then
Select Case Combo1.ListIndex
Case 0
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".suicune"
      Set m_pic = Nothing
      cmdNext.Enabled = False
      txtSourceImage.Text = ""
      'Suicune
      qAns = MsgBox("是否保留圖標背景色?", vbQuestion + vbYesNo, "Ask")
      If qAns = vbNo Then
      txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      Select Case Form1.Combo2.ListIndex
      '白色
      '蓝色
      '绿色
      '红色
      '黄色
      '黑色
      Case 0
      setTransparentColour vbWhite
      Case 1
      setTransparentColour vbBlue
      Case 2
      setTransparentColour vbGreen
      Case 3
      setTransparentColour vbRed
      Case 4
      setTransparentColour vbYellow
      Case 5
      setTransparentColour vbBlack
      End Select
      'setTransparentColour picTransparentColour.BackColor
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      Else
            txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      'setTransparentColour RGB(0, 245, 245)
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      End If
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
'MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
ans = MsgBox("文件已经存在,是否替换?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill bth & Text1.Text & ".jpg"
Select Case Combo1.ListIndex
Case 0
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".suicune"
      Set m_pic = Nothing
      cmdNext.Enabled = False
      txtSourceImage.Text = ""
      'Suicune
      qAns = MsgBox("是否保留圖標背景色?", vbQuestion + vbYesNo, "Ask")
      If qAns = vbNo Then
      txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      Select Case Form1.Combo2.ListIndex
      '白色
      '蓝色
      '绿色
      '红色
      '黄色
      '黑色
      Case 0
      setTransparentColour vbWhite
      Case 1
      setTransparentColour vbBlue
      Case 2
      setTransparentColour vbGreen
      Case 3
      setTransparentColour vbRed
      Case 4
      setTransparentColour vbYellow
      Case 5
      setTransparentColour vbBlack
      End Select
      'setTransparentColour picTransparentColour.BackColor
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      Else
            txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      'setTransparentColour RGB(0, 245, 245)
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      End If
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
'MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
Exit Sub
End If
End If
End If
End If
If Combo1.ListIndex = 2 Then
If Trim(Text1.Text) <> "" Then
If Dir(bth & Text1.Text & ".bmp") = "" Then
Select Case Combo1.ListIndex
Case 0
'SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".suicune"
      Set m_pic = Nothing
      cmdNext.Enabled = False
      txtSourceImage.Text = ""
      'Suicune
      qAns = MsgBox("是否保留圖標背景色?", vbQuestion + vbYesNo, "Ask")
      If qAns = vbNo Then
      txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      Select Case Form1.Combo2.ListIndex
      '白色
      '蓝色
      '绿色
      '红色
      '黄色
      '黑色
      Case 0
      setTransparentColour vbWhite
      Case 1
      setTransparentColour vbBlue
      Case 2
      setTransparentColour vbGreen
      Case 3
      setTransparentColour vbRed
      Case 4
      setTransparentColour vbYellow
      Case 5
      setTransparentColour vbBlack
      End Select
      'setTransparentColour picTransparentColour.BackColor
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      Else
            txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      'setTransparentColour RGB(0, 245, 245)
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      End If
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
'MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
ans = MsgBox("文件已经存在,是否替换?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill bth & Text1.Text & ".bmp"
Select Case Combo1.ListIndex
Case 0
'SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".suicune"
      Set m_pic = Nothing
      cmdNext.Enabled = False
      txtSourceImage.Text = ""
      'Suicune
      qAns = MsgBox("是否保留圖標背景色?", vbQuestion + vbYesNo, "Ask")
      If qAns = vbNo Then
      txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      Select Case Form1.Combo2.ListIndex
      '白色
      '蓝色
      '绿色
      '红色
      '黄色
      '黑色
      Case 0
      setTransparentColour vbWhite
      Case 1
      setTransparentColour vbBlue
      Case 2
      setTransparentColour vbGreen
      Case 3
      setTransparentColour vbRed
      Case 4
      setTransparentColour vbYellow
      Case 5
      setTransparentColour vbBlack
      End Select
      'setTransparentColour picTransparentColour.BackColor
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      Else
            txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      'setTransparentColour RGB(0, 245, 245)
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      End If
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
'MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
Exit Sub
End If
End If
End If
End If
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
End Sub
Private Sub Command2_Click()
On Error Resume Next
Unload Me
End Sub
Private Sub Dir1_Change()
On Error GoTo ep
Drive1.Drive = Left$(Dir1.Path, 2)
Select Case Combo1.ListIndex
Case 0
On Error Resume Next
With Me.WebBrowser1
.navigate Me.Dir1.Path
.Refresh
End With
With File1
.Pattern = "*.ico"
.Path = Dir1.Path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.TOp = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.navigate Dir1.Path
End With
Case 1
On Error Resume Next
With Me.WebBrowser1
.navigate Me.Dir1.Path
.Refresh
End With
With File1
.Pattern = "*.jpg"
.Path = Dir1.Path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.TOp = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.navigate Dir1.Path
End With
Case 2
On Error Resume Next
With Me.WebBrowser1
.navigate Me.Dir1.Path
.Refresh
End With
With File1
.Pattern = "*.bmp"
.Path = Dir1.Path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.TOp = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.navigate Dir1.Path
End With
End Select
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
On Error Resume Next
With Me.WebBrowser1
.navigate Me.Dir1.Path
.Refresh
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.TOp = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.navigate Dir1.Path
End With
End Sub
Private Sub Form_Activate()
On Error Resume Next
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.TOp = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.navigate Dir1.Path
End With
End Sub
Private Sub WebBrowser1_GotFocus()
On Error Resume Next
Me.Dir1.SetFocus
On Error Resume Next
Me.WebBrowser1.navigate "About:Operations Are Not Allowed "
End Sub
Private Sub Drive1_Change()
On Error GoTo ep
Dir1.Path = Drive1.Drive
Select Case Combo1.ListIndex
Case 0
On Error Resume Next
With Me.WebBrowser1
.navigate Me.Dir1.Path
.Refresh
End With
With File1
.Pattern = "*.ico"
.Path = Dir1.Path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.TOp = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.navigate Dir1.Path
End With
Case 1
On Error Resume Next
With Me.WebBrowser1
.navigate Me.Dir1.Path
.Refresh
End With
With File1
.Pattern = "*.jpg"
.Path = Dir1.Path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.TOp = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.navigate Dir1.Path
End With
Case 2
On Error Resume Next
With Me.WebBrowser1
.navigate Me.Dir1.Path
.Refresh
End With
With File1
.Pattern = "*.bmp"
.Path = Dir1.Path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.TOp = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.navigate Dir1.Path
End With
End Select
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
Drive1.Drive = "C:"
Dir1.Path = Drive1.Drive
Select Case Combo1.ListIndex
Case 0
With File1
.Pattern = "*.ico"
.Path = Dir1.Path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.TOp = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.navigate Dir1.Path
End With
Case 1
With File1
.Pattern = "*.jpg"
.Path = Dir1.Path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.TOp = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.navigate Dir1.Path
End With
Case 2
With File1
.Pattern = "*.bmp"
.Path = Dir1.Path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.TOp = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.navigate Dir1.Path
End With
End Select
End Sub
Private Sub File1_Click()
On Error GoTo ep
If File1.ListIndex >= 0 Then
Me.Text1.Text = File1.List(File1.ListIndex)
End If
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub File1_DblClick()
On Error GoTo ep
Dim ans As Integer
Me.Dir1.Enabled = False
Me.File1.Enabled = False
Me.Drive1.Enabled = False
Me.Command1.Enabled = False
Me.Command2.Enabled = False
Me.Combo1.Enabled = False
If Trim(Text1.Text) = "" Then
MsgBox "请输入文件名!", vbCritical, "Error"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
Exit Sub
End If
Select Case Right(Dir1.Path, 1)
Case "\"
bth = Dir1.Path
Case Else
bth = Dir1.Path & "\"
End Select
If Combo1.ListIndex = 0 Then
If Trim(Text1.Text) <> "" Then
If Dir(bth & Text1.Text & ".ico") = "" Then
Select Case Combo1.ListIndex
Case 0
'SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".suicune"
      Set m_pic = Nothing
      cmdNext.Enabled = False
      txtSourceImage.Text = ""
      'Suicune
      qAns = MsgBox("是否保留圖標背景色?", vbQuestion + vbYesNo, "Ask")
      If qAns = vbNo Then
      txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      Select Case Form1.Combo2.ListIndex
      '白色
      '蓝色
      '绿色
      '红色
      '黄色
      '黑色
      Case 0
      setTransparentColour vbWhite
      Case 1
      setTransparentColour vbBlue
      Case 2
      setTransparentColour vbGreen
      Case 3
      setTransparentColour vbRed
      Case 4
      setTransparentColour vbYellow
      Case 5
      setTransparentColour vbBlack
      End Select
      'setTransparentColour picTransparentColour.BackColor
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      Else
            txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      'setTransparentColour RGB(0, 245, 245)
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      End If
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
'MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
ans = MsgBox("文件已经存在,是否替换?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill bth & Text1.Text & ".ico"
Select Case Combo1.ListIndex
Case 0
'SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".suicune"
      Set m_pic = Nothing
      cmdNext.Enabled = False
      txtSourceImage.Text = ""
      'Suicune
      qAns = MsgBox("是否保留圖標背景色?", vbQuestion + vbYesNo, "Ask")
      If qAns = vbNo Then
      txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      Select Case Form1.Combo2.ListIndex
      '白色
      '蓝色
      '绿色
      '红色
      '黄色
      '黑色
      Case 0
      setTransparentColour vbWhite
      Case 1
      setTransparentColour vbBlue
      Case 2
      setTransparentColour vbGreen
      Case 3
      setTransparentColour vbRed
      Case 4
      setTransparentColour vbYellow
      Case 5
      setTransparentColour vbBlack
      End Select
      'setTransparentColour picTransparentColour.BackColor
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      Else
            txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      'setTransparentColour RGB(0, 245, 245)
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      End If
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
'MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
Exit Sub
End If
End If
End If
End If
If Combo1.ListIndex = 1 Then
If Trim(Text1.Text) <> "" Then
If Dir(bth & Text1.Text & ".jpg") = "" Then
Select Case Combo1.ListIndex
Case 0
'SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".suicune"
      Set m_pic = Nothing
      cmdNext.Enabled = False
      txtSourceImage.Text = ""
      'Suicune
      qAns = MsgBox("是否保留圖標背景色?", vbQuestion + vbYesNo, "Ask")
      If qAns = vbNo Then
      txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      Select Case Form1.Combo2.ListIndex
      '白色
      '蓝色
      '绿色
      '红色
      '黄色
      '黑色
      Case 0
      setTransparentColour vbWhite
      Case 1
      setTransparentColour vbBlue
      Case 2
      setTransparentColour vbGreen
      Case 3
      setTransparentColour vbRed
      Case 4
      setTransparentColour vbYellow
      Case 5
      setTransparentColour vbBlack
      End Select
      'setTransparentColour picTransparentColour.BackColor
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      Else
            txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      'setTransparentColour RGB(0, 245, 245)
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      End If
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
'MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
ans = MsgBox("文件已经存在,是否替换?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill bth & Text1.Text & ".jpg"
Select Case Combo1.ListIndex
Case 0
'SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".suicune"
      Set m_pic = Nothing
      cmdNext.Enabled = False
      txtSourceImage.Text = ""
      'Suicune
      qAns = MsgBox("是否保留圖標背景色?", vbQuestion + vbYesNo, "Ask")
      If qAns = vbNo Then
      txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      Select Case Form1.Combo2.ListIndex
      '白色
      '蓝色
      '绿色
      '红色
      '黄色
      '黑色
      Case 0
      setTransparentColour vbWhite
      Case 1
      setTransparentColour vbBlue
      Case 2
      setTransparentColour vbGreen
      Case 3
      setTransparentColour vbRed
      Case 4
      setTransparentColour vbYellow
      Case 5
      setTransparentColour vbBlack
      End Select
      'setTransparentColour picTransparentColour.BackColor
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      Else
            txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      'setTransparentColour RGB(0, 245, 245)
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      End If
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
'MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
Exit Sub
End If
End If
End If
End If
If Combo1.ListIndex = 2 Then
If Trim(Text1.Text) <> "" Then
If Dir(bth & Text1.Text & ".bmp") = "" Then
Select Case Combo1.ListIndex
Case 0
'SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".suicune"
      Set m_pic = Nothing
      cmdNext.Enabled = False
      txtSourceImage.Text = ""
      'Suicune
      qAns = MsgBox("是否保留圖標背景色?", vbQuestion + vbYesNo, "Ask")
      If qAns = vbNo Then
      txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      Select Case Form1.Combo2.ListIndex
      '白色
      '蓝色
      '绿色
      '红色
      '黄色
      '黑色
      Case 0
      setTransparentColour vbWhite
      Case 1
      setTransparentColour vbBlue
      Case 2
      setTransparentColour vbGreen
      Case 3
      setTransparentColour vbRed
      Case 4
      setTransparentColour vbYellow
      Case 5
      setTransparentColour vbBlack
      End Select
      'setTransparentColour picTransparentColour.BackColor
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      Else
            txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      'setTransparentColour RGB(0, 245, 245)
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      End If
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
'MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
ans = MsgBox("文件已经存在,是否替换?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill bth & Text1.Text & ".bmp"
Select Case Combo1.ListIndex
Case 0
'SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".suicune"
      Set m_pic = Nothing
      cmdNext.Enabled = False
      txtSourceImage.Text = ""
      'Suicune
      qAns = MsgBox("是否保留圖標背景色?", vbQuestion + vbYesNo, "Ask")
      If qAns = vbNo Then
      txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      Select Case Form1.Combo2.ListIndex
      '白色
      '蓝色
      '绿色
      '红色
      '黄色
      '黑色
      Case 0
      setTransparentColour vbWhite
      Case 1
      setTransparentColour vbBlue
      Case 2
      setTransparentColour vbGreen
      Case 3
      setTransparentColour vbRed
      Case 4
      setTransparentColour vbYellow
      Case 5
      setTransparentColour vbBlack
      End Select
      'setTransparentColour picTransparentColour.BackColor
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      Else
            txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      'setTransparentColour RGB(0, 245, 245)
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
            'Suicune
      If (fileExists(bth & Text1.Text & ".ico")) Then
         cFI.LoadIcon bth & Text1.Text & ".ico"
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         MsgBox "文件已經保存為" & vbCrLf & _
            bth & Text1.Text & ".ico", vbInformation, "Info"
            Kill bth & Text1.Text & ".suicune"
      Else
         MsgBox "保存圖標失敗", vbCritical, "Error"
      End If
      End If
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
'MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已经将图标保存为:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
Exit Sub
End If
End If
End If
End If
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
If KeyCode = vbKeyF5 Then
With File1
.Refresh
End With
End If
End Sub
Private Sub dir1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Debug.Print WebBrowser1.LocationURL
If UCase(Mid(WebBrowser1.LocationURL, 9, 1)) <> UCase(Left(Me.Drive1.Drive, 1)) Then
WebBrowser1.navigate Me.Dir1.Path
Else
Exit Sub
End If
End Sub
Private Sub file1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Debug.Print WebBrowser1.LocationURL
If UCase(Mid(WebBrowser1.LocationURL, 9, 1)) <> UCase(Left(Me.Drive1.Drive, 1)) Then
WebBrowser1.navigate Me.Dir1.Path
Else
Exit Sub
End If
End Sub
Private Sub text1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Debug.Print WebBrowser1.LocationURL
If UCase(Mid(WebBrowser1.LocationURL, 9, 1)) <> UCase(Left(Me.Drive1.Drive, 1)) Then
WebBrowser1.navigate Me.Dir1.Path
Else
Exit Sub
End If
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
Me.WebBrowser1.navigate "About:Operations Are Not Allowed "
End Sub
Private Sub Form_Load()
On Error Resume Next
   ' SPM - Form_Load is fired *after* the window is created.
   ' Therefore we will already have set the style so there
   ' is no need for the hook anymore.
   HookDetach
   
   SetIcon Me.hwnd, "AAA"
   
   picTab(1).Move picTab(0).Left, picTab(0).TOp, picTab(0).Width, picTab(0).Height
   picTab(2).Move picTab(0).Left, picTab(0).TOp, picTab(0).Width, picTab(0).Height
   picTab(0).BorderStyle = 0
   picTab(1).BorderStyle = 0
   picTab(2).BorderStyle = 0
   
Me.Combo1.ListIndex = 0
On Error Resume Next
With Me.WebBrowser1
.navigate Me.Dir1.Path
.Refresh
End With
With Me
.Left = Screen.Width / 2 - .Width / 2
.TOp = Screen.Height / 2 - .Height / 2
.Icon = LoadPicture()
.KeyPreview = True
End With
Text1.Text = ""
Select Case Combo1.ListIndex
Case 0
With File1
.Pattern = "*.ico"
.Path = Dir1.Path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
Case 1
With File1
.Pattern = "*.jpg"
.Path = Dir1.Path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
Case 0
With File1
.Pattern = "*.bmp"
.Path = Dir1.Path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
End Select
On Error Resume Next
With Me.WebBrowser1
.navigate Me.Dir1.Path
.Refresh
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.TOp = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.navigate Dir1.Path
End With
End Sub
Private Sub Drive1_GotFocus()
On Error Resume Next
Me.Command1.Default = False
Me.Command2.Cancel = False
End Sub
Private Sub Drive1_LostFocus()
On Error Resume Next
Me.Command1.Default = True
Me.Command2.Cancel = True
End Sub
Private Sub Dir1_GotFocus()
On Error Resume Next
Me.Command1.Default = False
Me.Command2.Cancel = True
End Sub
Private Sub Dir1_LostFocus()
On Error Resume Next
Me.Command1.Default = True
Me.Command2.Cancel = True
End Sub
Private Sub File1_GotFocus()
On Error Resume Next
Me.Command1.Default = True
Me.Command2.Cancel = True
End Sub
Private Sub File1_LostFocus()
On Error Resume Next
Me.Command1.Default = True
Me.Command2.Cancel = True
End Sub
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

   If (chkSize(0).Value = vbChecked) Then
      lIndex = cFI.IconIndex(16, 16, 32)
      If (lIndex = 0) Then
         lIndex = cFI.AddImage(16, 16, 32)
      End If
      createIconAtSize cFI, lIndex, 16, 16
   End If
   If (chkSize(1).Value = vbChecked) Then
      lIndex = cFI.IconIndex(24, 24, 32)
      If (lIndex = 0) Then
         lIndex = cFI.AddImage(24, 24, 32)
      End If
      createIconAtSize cFI, lIndex, 24, 24
   End If
   If (chkSize(2).Value = vbChecked) Then
      lIndex = cFI.IconIndex(32, 32, 32)
      If (lIndex = 0) Then
         lIndex = cFI.AddImage(32, 32, 32)
      End If
      createIconAtSize cFI, lIndex, 32, 32
   End If
   If (chkSize(3).Value = vbChecked) Then
      lIndex = cFI.IconIndex(48, 48, 32)
      If (lIndex = 0) Then
         lIndex = cFI.AddImage(48, 48, 32)
      End If
      createIconAtSize cFI, lIndex, 48, 48
   End If
   If (chkSize(4).Value = vbChecked) Then
   If 25 = 245 Then
      For i = 0 To lstCustom.ListCount
         sWidthHeight = lstCustom.List(i)
         iPos = InStr(sWidthHeight, "x")
         lWidth = CLng(Left(sWidthHeight, iPos - 1))
         lHeight = CLng(Mid(sWidthHeight, iPos + 1))
         lIndex = cFI.IconIndex(lWidth, lHeight, 32)
         If (lIndex = 0) Then
            lIndex = cFI.AddImage(lWidth, lHeight, 32)
         End If
         createIconAtSize cFI, lIndex, lWidth, lHeight
      Next i
      End If
            lIndex = cFI.IconIndex(64, 64, 32)
      If (lIndex = 0) Then
         lIndex = cFI.AddImage(64, 64, 32)
      End If
      createIconAtSize cFI, lIndex, 64, 64
   End If
      If (chkSize(5).Value = vbChecked) Then
   If 25 = 245 Then
      For i = 0 To lstCustom.ListCount
         sWidthHeight = lstCustom.List(i)
         iPos = InStr(sWidthHeight, "x")
         lWidth = CLng(Left(sWidthHeight, iPos - 1))
         lHeight = CLng(Mid(sWidthHeight, iPos + 1))
         lIndex = cFI.IconIndex(lWidth, lHeight, 32)
         If (lIndex = 0) Then
            lIndex = cFI.AddImage(lWidth, lHeight, 32)
         End If
         createIconAtSize cFI, lIndex, lWidth, lHeight
      Next i
      End If
            lIndex = cFI.IconIndex(128, 128, 32)
      If (lIndex = 0) Then
         lIndex = cFI.AddImage(128, 128, 32)
      End If
      createIconAtSize cFI, lIndex, 128, 128
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
   picSource.Cls
   m_cSource.AlphaPaintPicture picSource.hdc, _
      0, 0, _
      picSource.ScaleWidth \ Screen.TwipsPerPixelX, _
      picSource.ScaleHeight \ Screen.TwipsPerPixelY, _
      0, 0, _
      m_cSource.Width, m_cSource.Height
   picSource.Refresh
End Sub

Private Function fileExists(ByVal sFile As String) As Boolean
Dim sDir As String
   On Error Resume Next
   sDir = Dir(bth & Text1.Text & ".suicune")
   fileExists = (Len(sDir) > 0) And (Err.Number = 0)
End Function

Private Sub chkSize_Click(Index As Integer)
Dim bEnableNext As Boolean
Dim bEnableCustom As Boolean
Dim i As Long

   bEnableNext = (Len(txtFileName.Text) > 0)
   If (bEnableNext) Then
      For i = 0 To 3
         If (chkSize(i).Value = vbChecked) Then
            bEnableNext = True
            Exit For
         End If
      Next i
      If Not (bEnableNext) Then
         If (chkSize(4).Value = vbChecked) Then
            If (lstCustom.ListCount > 0) Then
               bEnableNext = True
            End If
         End If
      End If
   End If
   cmdNext.Enabled = bEnableNext
   
   bEnableCustom = (chkSize(4).Value = vbChecked)
   lstCustom.Enabled = bEnableCustom
   cmdAddCustom.Enabled = bEnableCustom
   cmdRemoveCustom.Enabled = ((lstCustom.ListIndex > -1) And bEnableCustom)
   
End Sub

Private Sub cmdAddCustom_Click()
Dim sR As String
Dim sWidth As String
Dim sHeight As String
Dim bValid As Boolean
Dim lWidth As Long
Dim lHeight As Long
Dim iPos As Long

   sR = InputBox("Enter custom icon size in the form [width] x [height]", App.Title, "")
   If Len(sR) > 0 Then
      iPos = InStr(sR, ",")
      If (iPos = 0) Then
         iPos = InStr(sR, "x")
      End If
      If (iPos > 0) Then
         If (iPos > 1) Then
            sWidth = Trim(Left(sR, iPos - 1))
            If (iPos < Len(sR)) Then
               sHeight = Trim(Mid(sR, iPos + 1))
               If (IsNumeric(sWidth) And IsNumeric(sHeight)) Then
                  On Error Resume Next
                  lWidth = CLng(sWidth)
                  If (Err.Number = 0) Then
                     lHeight = CLng(sHeight)
                     If (Err.Number = 0) Then
                        If (lWidth > 0) And (lHeight > 0) Then
                           If (lWidth = 16) And (lHeight = 16) Then
                           ElseIf (lWidth = 24) And (lHeight = 24) Then
                           ElseIf (lWidth = 32) And (lHeight = 32) Then
                           ElseIf (lWidth = 48) And (lHeight = 48) Then
                           Else
                              lstCustom.AddItem lWidth & "x" & lHeight
                              lstCustom.ListIndex = lstCustom.NewIndex
                              bValid = True
                           End If
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
      If Not (bValid) Then
         MsgBox "The custom icon size " & sR & " is not valid", vbInformation
      End If
   End If
End Sub

Private Sub cmdBack_Click()
   If (m_iTab = 2) Then
      picTab(1).Visible = True
      picTab(0).Visible = False
      picTab(2).Visible = False
      m_iTab = 1
      cmdBack.Enabled = True
      cmdNext.Enabled = (Len(txtFileName.Text) > 0)
   ElseIf (m_iTab = 1) Then
      picTab(0).Visible = True
      picTab(1).Visible = False
      picTab(2).Visible = False
      m_iTab = 0
      cmdBack.Enabled = False
      cmdNext.Enabled = Not (m_pic Is Nothing)
   End If
   
End Sub

Private Sub cmdCancel_Click()
   Unload Me
End Sub

Private Sub cmdNext_Click()
   If (m_iTab = 0) Then
      picTab(1).Visible = True
      picTab(0).Visible = False
      picTab(2).Visible = False
      cmdNext.Enabled = Len(txtFileName.Text) > 0
      cmdBack.Enabled = True
      cmdCancel.Caption = "取消(&C)"
      m_iTab = 1
   ElseIf (m_iTab = 1) Then
      picTab(2).Visible = True
      picTab(0).Visible = False
      picTab(1).Visible = False
      cmdNext.Enabled = False
      cmdBack.Enabled = True
      cmdStartOver.Enabled = False
      lblInfo.Caption = "正在創建圖標..."
      Me.Refresh
         
      'Suicune
      If (fileExists(txtFileName.Text)) Then
         cFI.LoadIcon txtFileName.Text
      End If
      createIcon cFI
      If cFI.SaveIcon(txtFileName.Text) Then
         lblInfo.Caption = "文件已經保存為" & vbCrLf & _
            txtFileName.Text
      Else
         lblInfo.Caption = "保存圖標失敗"
      End If
      cmdStartOver.Enabled = True
      cmdCancel.Caption = "取消(&C)"
      m_iTab = 2
   End If
End Sub

Private Sub cmdPickColour_Click()
Dim lColor As Long
Dim cD As New cCommonDialogLite
   OleTranslateColor picTransparentColour.BackColor, 0, lColor
   If (cD.VBChooseColor(lColor, FullOpen:=True, Owner:=Me.hwnd)) Then
      picTransparentColour.BackColor = lColor
      openImage
      setTransparentColour lColor
   End If
End Sub

Private Sub cmdPickInput_Click()
'Dim bth & Text1.Text & ".suicune" As String
Dim cD As New cCommonDialogLite
   'bth & Text1.Text & ".suicune" = txtSourceImage.Text
   If (cD.VBGetOpenFileName(bth & Text1.Text & ".suicune", _
      Filter:="Pictures(*.jpg;*.jpeg;*.bmp;*.bip;*.gif;*.wmf)|*.jpg;*.jpeg;*.bmp;*.bip;*.gif;*.wmf", _
      Owner:=Me.hwnd)) Then
      Set m_pic = Nothing
      cmdNext.Enabled = False
      txtSourceImage.Text = ""
      m_lTransparentColour = CLR_INVALID
      Set m_pic = LoadPicture(bth & Text1.Text & ".suicune")
      openImage
      picTransparentColour.BackColor = GetPixelAPI(m_cSource.hdc, 0, 0)
      setTransparentColour picTransparentColour.BackColor
      txtSourceImage.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = True
   End If
End Sub

Private Sub cmdPickOutput_Click()
'Dim bth & Text1.Text & ".suicune" As String
Dim cD As New cCommonDialogLite
   'bth & Text1.Text & ".suicune" = txtFileName.Text
   If (cD.VBGetSaveFileName(bth & Text1.Text & ".suicune", _
      Filter:="Icon Files (*.ICO)|*.ICO|All Files (*.*)|*.*", _
      DefaultExt:="ico", _
      Owner:=Me.hwnd)) Then
      txtFileName.Text = bth & Text1.Text & ".suicune"
      cmdNext.Enabled = Len(txtFileName.Text) > 0 And _
         ((chkSize(0).Value = vbChecked) Or _
         (chkSize(1).Value = vbChecked) Or _
         (chkSize(2).Value = vbChecked))
   End If
End Sub


Private Sub cmdRemoveCustom_Click()
Dim lIndex As Long
   If (lstCustom.ListIndex > -1) Then
      lIndex = lstCustom.ListIndex
      lstCustom.RemoveItem lstCustom.ListIndex
      If (lIndex < lstCustom.ListCount) Then
         lstCustom.ListIndex = lIndex
      Else
         If (lIndex - 1) > -1 Then
            lstCustom.ListIndex = lIndex - 1
         End If
      End If
      
      cmdRemoveCustom.Enabled = (lstCustom.ListIndex > -1)
   End If
End Sub

Private Sub cmdStartOver_Click()
   '
   cmdBack_Click
   cmdBack_Click
   '
End Sub

Private Sub Form_Initialize()

   ' SPM - Form_Initialize fired when object is being created,
   ' i.e. before hWnd created.
   HookAttach
   
   m_lTransparentColour = CLR_INVALID
   
End Sub



Private Sub lstCustom_Click()
   If (lstCustom.ListIndex > -1) Then
      cmdRemoveCustom.Enabled = True
   End If
End Sub


