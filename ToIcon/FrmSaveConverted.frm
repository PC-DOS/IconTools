VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FrmSaveConverted 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "保存图标"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13395
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   13395
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   5010
      Left            =   6960
      ScaleHeight     =   5010
      ScaleWidth      =   6390
      TabIndex        =   10
      Top             =   15
      Width           =   6390
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         CausesValidation=   0   'False
         Height          =   4965
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   6345
         ExtentX         =   11192
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
         Location        =   ""
      End
   End
   Begin VB.FileListBox File1 
      Height          =   4770
      Left            =   2865
      TabIndex        =   9
      Top             =   255
      Width           =   4080
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
      Height          =   315
      Left            =   675
      TabIndex        =   2
      Top             =   5040
      Width           =   12675
   End
   Begin VB.CommandButton Command1 
      Caption         =   "保存(&S)"
      Default         =   -1  'True
      Height          =   420
      Left            =   11025
      TabIndex        =   1
      Top             =   5400
      Width           =   1125
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   420
      Left            =   12210
      TabIndex        =   0
      Top             =   5400
      Width           =   1125
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
Attribute VB_Name = "FrmSaveConverted"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strPath As String
Dim strFile As String
Private Sub Form_Activate()
On Error Resume Next
With Me.WebBrowser1
.Navigate Me.Dir1.path
.Refresh
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.Top = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.Navigate Dir1.path
End With
End Sub
Private Sub Command1_Click()
On Error GoTo ep
Me.Command1.Enabled = False
Me.Command2.Enabled = False
Dir1.Enabled = False
Drive1.Enabled = False
File1.Enabled = False
If Trim(Text1.Text) <> "" Then
If Right(Dir1.path, 1) = "\" Then
strPath = Dir1.path
strFile = strPath & Text1.Text & ".ico"
If Dir(strFile) = "" Then
SavePicture Form2.Picture2.Image, strFile
MsgBox "已经将图标保存为:" & vbCrLf & strFile, vbExclamation, "Info"
Command1.Enabled = True
Command2.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
Unload Me
Exit Sub
Else
Dim ans As Integer
ans = MsgBox("文件已经存在,是否替换?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill strFile
SavePicture Form2.Picture2.Image, strFile
MsgBox "已经将图标保存为:" & vbCrLf & strFile, vbExclamation, "Info"
Command1.Enabled = True
Command2.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
Unload Me
Exit Sub
Else
Command1.Enabled = True
Command2.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
Exit Sub
End If
End If
Else
strPath = Dir1.path
strFile = strPath & "\" & Text1.Text & ".ico"
If Dir(strFile) = "" Then
SavePicture Form2.Picture2.Image, strFile
MsgBox "已经将图标保存为:" & vbCrLf & strFile, vbExclamation, "Info"
Command1.Enabled = True
Command2.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
Unload Me
Exit Sub
Else
ans = MsgBox("文件已经存在,是否替换?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill strFile
SavePicture Form2.Picture2.Image, strFile
MsgBox "已经将图标保存为:" & vbCrLf & strFile, vbExclamation, "Info"
Command1.Enabled = True
Command2.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
Unload Me
Exit Sub
Else
Command1.Enabled = True
Command2.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
Unload Me
Exit Sub
End If
End If
End If
Else
MsgBox "请输入文件名!", vbCritical, "Error"
Command1.Enabled = True
Command2.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
Unload Me
Exit Sub
End If
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
Command1.Enabled = True
Command2.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
End Sub
Private Sub Command2_Click()
On Error Resume Next
Unload Me
Form2.SetFocus
End Sub
Private Sub Dir1_Change()
On Error GoTo ep
Drive1.Drive = Left$(Dir1.path, 2)
With File1
.Pattern = "*.ico"
.path = Dir1.path
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
.Top = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.Navigate Dir1.path
End With
On Error Resume Next
With Me.WebBrowser1
.Navigate Me.Dir1.path
.Refresh
End With
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
On Error Resume Next
With Me.WebBrowser1
.Navigate Me.Dir1.path
.Refresh
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.Top = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.Navigate Dir1.path
End With
End Sub
Private Sub Drive1_Change()
On Error GoTo ep
Dir1.path = Drive1.Drive
With File1
.Pattern = "*.ico"
.path = Dir1.path
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
.Top = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.Navigate Dir1.path
End With
On Error Resume Next
With Me.WebBrowser1
.Navigate Me.Dir1.path
.Refresh
End With
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
Drive1.Drive = "C:"
Dir1.path = Drive1.Drive
With File1
.Pattern = "*.ico"
.path = Dir1.path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
On Error Resume Next
With Me.WebBrowser1
.Navigate Me.Dir1.path
.Refresh
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.Top = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.Navigate Dir1.path
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
Private Sub File1_DblClick()
On Error GoTo ep
Me.Command1.Enabled = False
Me.Command2.Enabled = False
Dir1.Enabled = False
Drive1.Enabled = False
File1.Enabled = False
If Trim(Text1.Text) <> "" Then
If Right(Dir1.path, 1) = "\" Then
strPath = Dir1.path
strFile = strPath & Text1.Text & ".ico"
If Dir(strFile) = "" Then
SavePicture Form2.Picture2.Image, strFile
MsgBox "已经将图标保存为:" & vbCrLf & strFile, vbExclamation, "Info"
Command1.Enabled = True
Command2.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
Unload Me
Exit Sub
Else
Dim ans As Integer
ans = MsgBox("文件已经存在,是否替换?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill strFile
SavePicture Form2.Picture2.Image, strFile
MsgBox "已经将图标保存为:" & vbCrLf & strFile, vbExclamation, "Info"
Command1.Enabled = True
Command2.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
Unload Me
Exit Sub
Else
Command1.Enabled = True
Command2.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
Exit Sub
End If
End If
Else
strPath = Dir1.path
strFile = strPath & "\" & Text1.Text & ".ico"
If Dir(strFile) = "" Then
SavePicture Form2.Picture2.Image, strFile
MsgBox "已经将图标保存为:" & vbCrLf & strFile, vbExclamation, "Info"
Command1.Enabled = True
Command2.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
Unload Me
Exit Sub
Else
ans = MsgBox("文件已经存在,是否替换?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill strFile
SavePicture Form2.Picture2.Image, strFile
MsgBox "已经将图标保存为:" & vbCrLf & strFile, vbExclamation, "Info"
Command1.Enabled = True
Command2.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
Unload Me
Exit Sub
Else
Command1.Enabled = True
Command2.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
Unload Me
Exit Sub
End If
End If
End If
Else
MsgBox "请输入文件名!", vbCritical, "Error"
Command1.Enabled = True
Command2.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
Unload Me
Exit Sub
End If
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
Command1.Enabled = True
Command2.Enabled = True
Dir1.Enabled = True
Drive1.Enabled = True
File1.Enabled = True
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
Private Sub File1_Click()
On Error GoTo ep
If File1.ListIndex >= 0 Then
Me.Text1.Text = File1.List(File1.ListIndex)
End If
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
End Sub
Private Sub Form_Load()
On Error Resume Next
Me.Command1.Default = True
Me.Command2.Cancel = True
With Me
.Left = Screen.Width / 2 - .Width / 2
.Top = Screen.Height / 2 - .Height / 2
.Icon = LoadPicture()
.KeyPreview = True
End With
With File1
.path = Me.Dir1.path
.Pattern = "*.ico"
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
Me.Text1.Text = ""
On Error Resume Next
With Me.WebBrowser1
.Navigate Me.Dir1.path
.Refresh
End With
With Me.Picture1
.Enabled = False
.BorderStyle = 0
End With
With Me.WebBrowser1
.Left = 0
.Top = 0
.Height = Me.Picture1.ScaleHeight
.Width = Me.Picture1.ScaleWidth
.Navigate Dir1.path
End With
End Sub
Private Sub text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Debug.Print WebBrowser1.LocationURL
If UCase(Mid(WebBrowser1.LocationURL, 9, 1)) <> UCase(Left(Me.Drive1.Drive, 1)) Then
WebBrowser1.Navigate Me.Dir1.path
Else
Exit Sub
End If
End Sub
Private Sub WebBrowser1_GotFocus()
On Error Resume Next
Me.Dir1.SetFocus
On Error Resume Next
Me.WebBrowser1.Navigate "About:Operations Are Not Allowed "
End Sub
Private Sub dir1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Debug.Print WebBrowser1.LocationURL
If UCase(Mid(WebBrowser1.LocationURL, 9, 1)) <> UCase(Left(Me.Drive1.Drive, 1)) Then
WebBrowser1.Navigate Me.Dir1.path
Else
Exit Sub
End If
End Sub
Private Sub file1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Debug.Print WebBrowser1.LocationURL
If UCase(Mid(WebBrowser1.LocationURL, 9, 1)) <> UCase(Left(Me.Drive1.Drive, 1)) Then
WebBrowser1.Navigate Me.Dir1.path
Else
Exit Sub
End If
End Sub
Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
Me.WebBrowser1.Navigate "About:Operations Are Not Allowed "
End Sub
