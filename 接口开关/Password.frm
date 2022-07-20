VERSION 5.00
Begin VB.Form Password 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "请输入密码"
   ClientHeight    =   3255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Password.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   2895
   StartUpPosition =   1  '所有者中心
   Begin VB.TextBox Text4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   5
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1560
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   4
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   960
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   3
      Top             =   720
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   360
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   2
      Top             =   720
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   "小键盘"
      Height          =   1815
      Left            =   360
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
      Begin VB.CommandButton CommandNum 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton CommandNum 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   9
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton CommandNum 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton CommandNum 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton CommandNum 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   6
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton CommandNum 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton CommandNum 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton CommandNum 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   3
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton CommandNum 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   600
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.CommandButton CommandNum 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Label LabelInt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "随机码：000   时间：00:00"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   360
      TabIndex        =   16
      Top             =   120
      Width           =   2280
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "键入动态密码以解锁功能——"
      Height          =   195
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2280
   End
End
Attribute VB_Name = "Password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Dim It As Integer
Dim It1 As Integer
Dim It2 As Integer
Dim It3 As Integer
Dim Hr As Integer
Dim Hr1 As Integer
Dim Hr2 As Integer
Dim Mn As Integer
Dim Mn1 As Integer
Dim Mn2 As Integer
Dim Pwd1 As Integer
Dim Pwd2 As Integer
Dim Pwd3 As Integer
Dim Pwd4 As Integer
'在计算密码时可供使用的变量：
'It     随机码，取值100~999的整数
'It1     随机码百位，取值1~9的整数
'It2     随机码十位，取值0~9的整数
'It3     随机码个位，取值0~9的整数
'Hr     系统小时数，0~23的整数
'Hr1    小时数十位，0~2的整数
'Hr2    小时数个位，0~9的整数
'Mn     分钟数，0~59的整数
'Mn1    分钟数十位，0~5的整数
'Mn2    分钟数个位，0~9的整数

Function Pwd1True(Pwd As Integer) As Boolean
Dim TruePwd As Integer
TruePwd = (It1 + Hr1) Mod 10                                      '第一位算法
'我的设置为：随机码百位+小时数十位，结果取各位（除以10取余不就是取个位么）
'以下密码第二位什么的类推吧
If Pwd = TruePwd Then
Pwd1True = True
Else
Pwd1True = False
End If
End Function

Function Pwd2True(Pwd As Integer) As Boolean
Dim TruePwd As Integer
TruePwd = Abs((It1 - Hr1) Mod 10)                                 '第二位算法
If Pwd = TruePwd Then
Pwd2True = True
Else
Pwd2True = False
End If
End Function

Function Pwd3True(Pwd As Integer) As Boolean
Dim TruePwd As Integer
TruePwd = (It3 + Mn2) Mod 10                                      '第三位算法
If Pwd = TruePwd Then
Pwd3True = True
Else
Pwd3True = False
End If
End Function

Function Pwd4True(Pwd As Integer) As Boolean
Dim TruePwd As Integer
TruePwd = Abs((It3 - Mn2) Mod 10)                                 '第四位算法
If Pwd = TruePwd Then
Pwd4True = True
Else
Pwd4True = False
End If
End Function

Private Sub Form_Load()
SkinH_Attach
SkinH.SkinH_SetAero 1
Randomize   '初始化随机数
It = Int((900 * Rnd) + 100)
It1 = It \ 100
It2 = It \ 10 Mod 10
It3 = It Mod 10
Hr = Hour(Now)
Hr1 = Hr \ 10
Hr2 = Hr Mod 10
Mn = Minute(Now)
Mn1 = Mn \ 10
Mn2 = Mn Mod 10
LabelInt.Caption = "随机码：" & It & "   时间：" & Format(Hr, "00") & ":" & Format(Mn, "00")
End Sub

Private Sub CommandNum_Click(Index As Integer)
If Text1.Text = "" Then
Text1.Text = Index
Text2.SetFocus
Else
  If Text2.Text = "" Then
  Text2.Text = Index
  Text3.SetFocus
  Else
    If Text3.Text = "" Then
    Text3.Text = Index
    Text4.SetFocus
    Else
      Text4.Text = Index
      Pwd1 = Text1.Text
      Pwd2 = Text2.Text
      Pwd3 = Text3.Text
      Pwd4 = Text4.Text
      Sleep 100
      If Pwd1True(Pwd1) = True And Pwd2True(Pwd2) = True And Pwd3True(Pwd3) = True And Pwd4True(Pwd4) = True Then
      Label1.Caption = "密码输入正确"
      '↓↓↓密码输入正确时干啥↓↓↓
      Shell "C:\PROGRA~1\VCOMTO~1\开启接口V1.01.exe"
      If MsgBox("现在重启以应用本次对电脑的所有更改？", vbYesNo, "写保护已开启") = vbYes Then
      Shell "cmd.exe /c shutdown -r -t 0"
      End If
      End
      '↑↑↑密码输入正确时干啥↑↑↑
      Else
      Label1.Caption = "密码输入错误"
      Text1.Text = ""
      Text2.Text = ""
      Text3.Text = ""
      Text4.Text = ""
      Text1.SetFocus
      End If
    End If
  End If
End If
End Sub
