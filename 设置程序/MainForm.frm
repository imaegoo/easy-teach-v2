VERSION 5.00
Begin VB.Form MainForm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "班班通助手 设置向导  Free ＆ Fast ＆ Simple"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7905
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   527
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      Caption         =   "检查更新"
      Height          =   360
      Left            =   6120
      TabIndex        =   5
      Top             =   4920
      Width           =   1410
   End
   Begin VB.CommandButton Command4 
      Caption         =   "测试版源代码公开下载 (VB6)"
      Height          =   360
      Left            =   360
      TabIndex        =   4
      Top             =   4920
      Width           =   3090
   End
   Begin VB.CommandButton Command3 
      Caption         =   "调整其它设置？点我！"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   1920
      TabIndex        =   2
      Top             =   3600
      Width           =   4050
   End
   Begin VB.CommandButton Command2 
      Caption         =   "调整课程表？点我！"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   1920
      TabIndex        =   1
      Top             =   2400
      Width           =   4050
   End
   Begin VB.CommandButton Command1 
      Caption         =   "第一次使用？点我！"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   840
      Left            =   1920
      TabIndex        =   0
      Top             =   1200
      Width           =   4050
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "欢迎回来，如果白板启动有异常，点击第一次使用重新校准"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00D9A108&
      Height          =   330
      Left            =   840
      TabIndex        =   3
      Top             =   480
      Width           =   6240
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Command1_Click()
Step1.Show 0
Unload Me
End Sub

Private Sub Command2_Click()
Step2.Show 0
Unload Me
End Sub

Private Sub Command3_Click()
Step3.Show 0
Unload Me
End Sub

Private Sub Command4_Click()
r = ShellExecute(0, "open", "http://mae.wodemo.com/", 0, 0, 1)
End Sub

Private Sub Command5_Click()
Shell App.Path & "\Update.exe", 1
End Sub

Private Sub Form_Load()
SkinH_Attach
SkinH.SkinH_SetAero 1
If Dir(App.Path & "\Config.txt") = "" Then
Label1.ForeColor = &HFF&
Label1.Caption = "未检测到配置文件，请点击“第一次使用”"
Command2.Enabled = False
Command3.Enabled = False
End If
End Sub
