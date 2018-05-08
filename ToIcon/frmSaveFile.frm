VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
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
      ItemData        =   "frmSaveFile.frx":0000
      Left            =   945
      List            =   "frmSaveFile.frx":000D
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
      Cancel          =   -1  'True
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
      Caption         =   "文件類型"
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
      Caption         =   "驅動器"
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
MsgBox "請輸入文件名!", vbCritical, "Error"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
Exit Sub
End If
Select Case Right(Dir1.path, 1)
Case "\"
bth = Dir1.path
Case Else
bth = Dir1.path & "\"
End Select
If Combo1.ListIndex = 0 Then
If Trim(Text1.Text) <> "" Then
If Dir(bth & Text1.Text & ".ico") = "" Then
Select Case Combo1.ListIndex
Case 0
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
ans = MsgBox("文件已經存在，是否替換?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill bth & Text1.Text & ".ico"
Select Case Combo1.ListIndex
Case 0
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
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
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
ans = MsgBox("文件已經存在，是否替換?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill bth & Text1.Text & ".jpg"
Select Case Combo1.ListIndex
Case 0
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
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
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
ans = MsgBox("文件已經存在，是否替換?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill bth & Text1.Text & ".bmp"
Select Case Combo1.ListIndex
Case 0
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
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
Drive1.Drive = Left$(Dir1.path, 2)
Select Case Combo1.ListIndex
Case 0
On Error Resume Next
With Me.WebBrowser1
.Navigate Me.Dir1.path
.Refresh
End With
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
Case 1
On Error Resume Next
With Me.WebBrowser1
.Navigate Me.Dir1.path
.Refresh
End With
With File1
.Pattern = "*.jpg"
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
Case 2
On Error Resume Next
With Me.WebBrowser1
.Navigate Me.Dir1.path
.Refresh
End With
With File1
.Pattern = "*.bmp"
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
End Select
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
Private Sub Form_Activate()
On Error Resume Next
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
Private Sub WebBrowser1_GotFocus()
On Error Resume Next
Me.Dir1.SetFocus
On Error Resume Next
Me.WebBrowser1.Navigate "About:Operations Are Not Allowed "
End Sub
Private Sub Drive1_Change()
On Error GoTo ep
Dir1.path = Drive1.Drive
Select Case Combo1.ListIndex
Case 0
On Error Resume Next
With Me.WebBrowser1
.Navigate Me.Dir1.path
.Refresh
End With
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
Case 1
On Error Resume Next
With Me.WebBrowser1
.Navigate Me.Dir1.path
.Refresh
End With
With File1
.Pattern = "*.jpg"
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
Case 2
On Error Resume Next
With Me.WebBrowser1
.Navigate Me.Dir1.path
.Refresh
End With
With File1
.Pattern = "*.bmp"
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
End Select
Exit Sub
ep:
MsgBox Err.Description, vbCritical, "Error"
Drive1.Drive = "C:"
Dir1.path = Drive1.Drive
Select Case Combo1.ListIndex
Case 0
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
Case 1
With File1
.Pattern = "*.jpg"
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
Case 2
With File1
.Pattern = "*.bmp"
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
MsgBox "請輸入文件名!", vbCritical, "Error"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
Exit Sub
End If
Select Case Right(Dir1.path, 1)
Case "\"
bth = Dir1.path
Case Else
bth = Dir1.path & "\"
End Select
If Combo1.ListIndex = 0 Then
If Trim(Text1.Text) <> "" Then
If Dir(bth & Text1.Text & ".ico") = "" Then
Select Case Combo1.ListIndex
Case 0
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
ans = MsgBox("文件已經存在，是否替換?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill bth & Text1.Text & ".ico"
Select Case Combo1.ListIndex
Case 0
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
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
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
ans = MsgBox("文件已經存在，是否替換?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill bth & Text1.Text & ".jpg"
Select Case Combo1.ListIndex
Case 0
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
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
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
Unload Me
End Select
Else
ans = MsgBox("文件已經存在，是否替換?", vbExclamation + vbYesNo, "Ask")
If ans = vbYes Then
Kill bth & Text1.Text & ".bmp"
Select Case Combo1.ListIndex
Case 0
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".ico"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".ico", vbExclamation, "Info"
Unload Me
Case 1
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".jpg"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".jpg", vbExclamation, "Info"
Unload Me
Case 2
SavePicture Form1.Picture3.Image, bth & Text1.Text & ".bmp"
Me.Dir1.Enabled = True
Me.File1.Enabled = True
Me.Drive1.Enabled = True
Me.Command1.Enabled = True
Me.Command2.Enabled = True
Me.Combo1.Enabled = True
MsgBox "已經將圖標保存為:" & vbCrLf & bth & Text1.Text & ".bmp", vbExclamation, "Info"
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
Private Sub text1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
Private Sub Form_Load()
Me.Combo1.ListIndex = 0
On Error Resume Next
With Me.WebBrowser1
.Navigate Me.Dir1.path
.Refresh
End With
With Me
.Left = Screen.Width / 2 - .Width / 2
.Top = Screen.Height / 2 - .Height / 2
.Icon = LoadPicture()
.KeyPreview = True
End With
Text1.Text = ""
Select Case Combo1.ListIndex
Case 0
With File1
.Pattern = "*.ico"
.path = Dir1.path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
Case 1
With File1
.Pattern = "*.jpg"
.path = Dir1.path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
Case 0
With File1
.Pattern = "*.bmp"
.path = Dir1.path
.System = True
.Hidden = True
.ReadOnly = True
.Normal = True
.Archive = True
End With
End Select
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
