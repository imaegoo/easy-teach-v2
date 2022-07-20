VERSION 5.00
Begin VB.Form Step1 
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
   Icon            =   "Step1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   527
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   4920
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下一步"
      Enabled         =   0   'False
      Height          =   360
      Left            =   6600
      TabIndex        =   11
      Top             =   4920
      Width           =   1005
   End
   Begin VB.Frame FraStep1 
      Caption         =   "第一步 智能检测"
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7455
      Begin VB.PictureBox PictureJia 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   435
         Left            =   5040
         Picture         =   "Step1.frx":4322
         ScaleHeight     =   435
         ScaleWidth      =   450
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1560
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.CommandButton cmdAuto 
         Caption         =   "开始检测"
         Height          =   360
         Left            =   6360
         TabIndex        =   1
         Top             =   240
         Width           =   990
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "写入初始配置文件"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00946E50&
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   3000
         Width           =   2280
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002E2DC3&
         Height          =   1080
         Left            =   2520
         TabIndex        =   9
         Top             =   2880
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "获取屏幕白板坐标信息"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00946E50&
         Height          =   375
         Left            =   2280
         TabIndex        =   8
         Top             =   2280
         Width           =   2850
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "启动白板主程序"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00946E50&
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   1560
         Width           =   1995
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "班班通系统认证"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00946E50&
         Height          =   375
         Left            =   720
         TabIndex        =   6
         Top             =   960
         Width           =   1995
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002E2DC3&
         Height          =   1080
         Left            =   1680
         TabIndex        =   5
         Top             =   2160
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002E2DC3&
         Height          =   1080
         Left            =   960
         TabIndex        =   4
         Top             =   1440
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   48
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002E2DC3&
         Height          =   1080
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   540
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "请点击“开始检测”"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   1380
      End
   End
End
Attribute VB_Name = "Step1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
'获取屏幕某点颜色
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
'获取屏幕某点颜色相关
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'获取鼠标位置
Private Type POINTAPI    '创建用户自定义的类型
    x As Long    '定义X的数据类型
    y As Long    '定义Y的数据类型
End Type
Dim Djs20 As Integer

Private Sub Command1_Click()
Step2.Show 0
Unload Me
End Sub

Private Sub Form_Load()
SkinH_Attach
SkinH.SkinH_SetAero 1
Djs20 = 21
If Dir(App.Path & "\Config.txt") <> "" Then Command1.Enabled = True
End Sub

Private Sub cmdAuto_Click()

    If cmdAuto.Caption = "取消检测" Then
        Timer1.Enabled = False
        cmdAuto.Caption = "开始检测"
    Else
        cmdAuto.Caption = "取消检测"
        If Dir("C:\Program Files\HiteBoard\HiteBoard\") <> "" Then
            Label3.ForeColor = &HFF00& '大绿
            Label6.ForeColor = &HFF00&
            Label6.Caption = "班班通系统认证　已通过"
        Else
            Label3.ForeColor = &HFF& '大红
            Label6.ForeColor = &HFF&
            MsgBox "检测到您的系统非九中班班通，继续运行会出现兼容性问题，自动检测已中止", 48, "警告"
            Exit Sub
        End If

        Shell "C:\Program Files\HiteBoard\HiteBoard\Environment.exe"
        Label4.ForeColor = &HAB8FE '土黄
        Timer1.Enabled = True
    End If

End Sub

Private Sub Timer1_Timer()
    Label7.Caption = "白板工具栏出现后，点击“      ”一次"
    PictureJia.Visible = True

    If GetPixel(GetWindowDC(0), 511, 384) = 16777215 Then
        Dim P1 As POINTAPI
        GetCursorPos P1
        Label4.ForeColor = &HFF00& '大绿
        Label7.ForeColor = &HFF00&
        Label8.Caption = "成功获取启动坐标 (" & P1.x & ", " & P1.y & ")"
        Label5.ForeColor = &HFF00& '大绿
        Label8.ForeColor = &HFF00&
        Shell "taskkill /F /im Environment.exe"
        Open App.Path & "\Config.txt" For Output As #3
        Write #3, "[白板坐标]"
        Write #3, P1.x
        Write #3, P1.y
        Close #3
        Label9.ForeColor = &HFF00& '大绿
        Label12.ForeColor = &HFF00&
        Label12.Caption = "成功写入，单击下一步继续"
        cmdAuto.Caption = "重新检测"
        Command1.Enabled = True
        Timer1.Enabled = False
    End If

End Sub
