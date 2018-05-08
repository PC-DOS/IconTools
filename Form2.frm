VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pic2Ico"
   ClientHeight    =   6180
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10440
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6180
   ScaleWidth      =   10440
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "关闭(&C)"
      Height          =   1260
      Left            =   6945
      Picture         =   "Form2.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4875
      Width           =   3390
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存为ICO文件(&S)"
      Height          =   1260
      Left            =   6945
      Picture         =   "Form2.frx":0442
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2835
      Width           =   3390
   End
   Begin VB.Frame Frame2 
      Caption         =   "转换"
      Height          =   3450
      Left            =   30
      TabIndex        =   4
      Top             =   2715
      Width           =   6810
      Begin VB.PictureBox Picture1 
         Height          =   3150
         Left            =   75
         ScaleHeight     =   206
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   259
         TabIndex        =   11
         Top             =   180
         Width           =   3945
         Begin VB.PictureBox Picture2 
            Height          =   300
            Left            =   1950
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   12
            Top             =   1200
            Width           =   300
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "大小"
         Enabled         =   0   'False
         Height          =   3180
         Left            =   4095
         TabIndex        =   5
         Top             =   165
         Width           =   2595
         Begin VB.OptionButton Option5 
            Caption         =   "&128*128"
            Enabled         =   0   'False
            Height          =   405
            Left            =   150
            TabIndex        =   10
            Top             =   2505
            Width           =   2190
         End
         Begin VB.OptionButton Option4 
            Caption         =   "&64*64"
            Enabled         =   0   'False
            Height          =   405
            Left            =   150
            TabIndex        =   9
            Top             =   1965
            Width           =   2190
         End
         Begin VB.OptionButton Option3 
            Caption         =   "&48*48"
            Enabled         =   0   'False
            Height          =   405
            Left            =   150
            TabIndex        =   8
            Top             =   1425
            Width           =   2190
         End
         Begin VB.OptionButton Option2 
            Caption         =   "&32*32"
            Enabled         =   0   'False
            Height          =   405
            Left            =   150
            TabIndex        =   7
            Top             =   900
            Value           =   -1  'True
            Width           =   2190
         End
         Begin VB.OptionButton Option1 
            Caption         =   "&16*16"
            Enabled         =   0   'False
            Height          =   405
            Left            =   150
            TabIndex        =   6
            Top             =   360
            Width           =   2190
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "浏览图片"
      Height          =   2715
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   10410
      Begin VB.DriveListBox Drive1 
         Height          =   300
         Left            =   135
         TabIndex        =   3
         Top             =   270
         Width           =   3765
      End
      Begin VB.DirListBox Dir1 
         Height          =   1980
         Left            =   135
         TabIndex        =   2
         Top             =   660
         Width           =   3750
      End
      Begin VB.FileListBox File1 
         Height          =   2430
         Hidden          =   -1  'True
         Left            =   3915
         Pattern         =   "*.bmp;*.jpg;*.jpeg;*.gif;*.wmf"
         System          =   -1  'True
         TabIndex        =   1
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         Caption         =   "图片预览窗格"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7650
         TabIndex        =   15
         Top             =   1275
         Width           =   1890
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   2370
         Left            =   6870
         Stretch         =   -1  'True
         Top             =   270
         Width           =   3435
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
On Error Resume Next
If Image1.Picture = LoadPicture() Then
MsgBox "请事先选择一个有效的图片文件!", vbCritical, "Error"
Command1.Enabled = False
Exit Sub
Else
FrmSaveConverted.Show 1
End If
End Sub
Private Sub Command2_Click()
On Error Resume Next
Unload Me
Form1.SetFocus
End Sub
Private Sub Form_Load()
On Error Resume Next
File1.ListIndex = -1
Image1.Picture = LoadPicture()
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
Drive1.Drive = "C:"
Image1.Picture = LoadPicture()
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.Top = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.Top = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
File1.ListIndex = -1
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
With Me.Command2
.Enabled = True
.Cancel = True
End With
With Me.Command1
.Enabled = False
.Default = False
End With
With Me.Label1
.Enabled = False
.AutoSize = True
.Visible = True
End With
With Me
.Icon = LoadPicture()
.Left = Screen.Width / 2 - .Width / 2
.Top = Screen.Height / 2 - .Height / 2
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.Top = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.Top = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
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
Private Sub Dir1_Change()
On Error GoTo ep
With File1
.Path = Dir1.Path
.Pattern = "*.bmp;*.jpg;*.jpeg;*.wmf;*.gif"
End With
File1.ListIndex = -1
Image1.Picture = LoadPicture()
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
Image1.Picture = LoadPicture()
File1.ListIndex = -1
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.Top = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.Top = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
With Me.Command2
.Enabled = True
.Cancel = True
End With
With Me.Command1
.Enabled = False
.Default = False
End With
With Me.Label1
.Enabled = False
.AutoSize = True
.Visible = True
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.Top = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.Top = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
Exit Sub
ep:
MsgBox "发生意外错误:" & vbCrLf & Err.Description, vbCritical, "Error"
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.Top = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.Top = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
File1.ListIndex = -1
Image1.Picture = LoadPicture()
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
Image1.Picture = LoadPicture()
File1.ListIndex = -1
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
With Me.Command2
.Enabled = True
.Cancel = True
End With
With Me.Command1
.Enabled = False
.Default = False
End With
With Me.Label1
.Enabled = False
.AutoSize = True
.Visible = True
End With
End Sub
Private Sub Drive1_Change()
On Error GoTo ep
Dir1.Path = Drive1.Drive
File1.Path = Dir1.Path
File1.ListIndex = -1
Image1.Picture = LoadPicture()
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
Image1.Picture = LoadPicture()
File1.ListIndex = -1
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.Top = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.Top = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
With Me.Command2
.Enabled = True
.Cancel = True
End With
With Me.Command1
.Enabled = False
.Default = False
End With
With Me.Label1
.Enabled = False
.AutoSize = True
.Visible = True
End With
Exit Sub
ep:
MsgBox "发生意外错误:" & vbCrLf & Err.Description, vbCritical, "Error"
File1.ListIndex = -1
Image1.Picture = LoadPicture()
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
Drive1.Drive = "C:"
Image1.Picture = LoadPicture()
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.Top = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.Top = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.Refresh
End With
File1.ListIndex = -1
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
With Me.Command2
.Enabled = True
.Cancel = True
End With
With Me.Command1
.Enabled = False
.Default = False
End With
With Me.Label1
.Enabled = False
.AutoSize = True
.Visible = True
End With
End Sub
Private Sub File1_Click()
On Error GoTo ep
If File1.ListIndex >= 0 Then
If Right(Dir1.Path, 1) = "\" Then
Image1.Picture = LoadPicture()
Image1.Picture = LoadPicture(Dir1.Path & File1.List(File1.ListIndex))
Option5.Enabled = True
Option4.Enabled = True
Option3.Enabled = True
Option2.Enabled = True
Option1.Enabled = True
Frame3.Enabled = True
Option2.Value = True
Command1.Enabled = True
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.Top = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.Top = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 32, 32
.Refresh
End With
With Me.Label1
.Enabled = False
.AutoSize = True
.Visible = False
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
Image1.Picture = LoadPicture()
Image1.Picture = LoadPicture(Dir1.Path & "\" & File1.List(File1.ListIndex))
Option5.Enabled = True
Option4.Enabled = True
Option3.Enabled = True
Option2.Enabled = True
Option1.Enabled = True
Command1.Enabled = True
Frame3.Enabled = True
Option2.Value = True
With Me.Label1
.Enabled = False
.AutoSize = True
.Visible = False
End With
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.Top = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.Top = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
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
Else
Exit Sub
End If
Exit Sub
ep:
MsgBox "发生意外错误:" & vbCrLf & Err.Description, vbCritical, "Error"
Image1.Picture = LoadPicture()
File1.ListIndex = -1
Option5.Enabled = False
Option4.Enabled = False
Option3.Enabled = False
Option2.Enabled = False
Option1.Enabled = False
Frame3.Enabled = False
Command1.Enabled = False
With Me.Command2
.Enabled = True
.Cancel = True
End With
With Me.Command1
.Enabled = False
.Default = False
End With
With Me.Label1
.Enabled = False
.AutoSize = True
.Visible = True
End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload Me
Form1.SetFocus
End Sub
Private Sub Option1_Click()
On Error GoTo ep
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 20
.Width = 20
.Left = 0
.Top = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.Top = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 16, 16
.Refresh
End With
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub Option2_Click()
On Error GoTo ep
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 36
.Width = 36
.Left = 0
.Top = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.Top = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 32, 32
.Refresh
End With
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub Option3_Click()
On Error GoTo ep
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 52
.Width = 52
.Left = 0
.Top = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.Top = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 48, 48
.Refresh
End With
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub Option4_Click()
On Error GoTo ep
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 68
.Width = 68
.Left = 0
.Top = 0
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.Top = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 64, 64
.Refresh
End With
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub Option5_Click()
On Error GoTo ep
With Me.Picture2
.AutoRedraw = True
.Cls
.ScaleMode = 3
.Height = 132
.Width = 132
.Left = Picture1.Width / Screen.TwipsPerPixelX / 2 - .Width / 2
.Top = Picture1.Height / Screen.TwipsPerPixelY / 2 - .Height / 2
.PaintPicture Me.Image1.Picture, 0, 0, 128, 128
.Refresh
End With
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub Picture2_Click()
On Error GoTo ep
If Me.Image1.Picture <> LoadPicture() Then
With Clipboard
.Clear
.SetData Me.Picture2.Image
End With
MsgBox "已经将图标发送到剪切板,可以在任何绘图软件中使用CTRL+V或其它粘贴快捷键粘贴!", vbExclamation, "Info"
Exit Sub
Exit Sub
Else
Exit Sub
End If
Exit Sub
ep:
MsgBox "复制图片失败,错误" & Err.Number & vbCrLf & Err.Description, vbCritical, "Error"
Exit Sub
Exit Sub
End Sub
