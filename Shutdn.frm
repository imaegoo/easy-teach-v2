VERSION 5.00
Begin VB.Form Shutdn 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "一键关机"
   ClientHeight    =   1650
   ClientLeft      =   8340
   ClientTop       =   8295
   ClientWidth     =   3075
   Icon            =   "Shutdn.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   906.338
   ScaleMode       =   0  'User
   ScaleWidth      =   1598.694
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   120
      Top             =   120
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H00FFFFFF&
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H008080FF&
      Caption         =   "立即关机"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   1815
      Left            =   0
      Picture         =   "Shutdn.frx":4322
      ScaleHeight     =   1755
      ScaleWidth      =   3075
      TabIndex        =   2
      Top             =   0
      Width           =   3135
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   26.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   840
         TabIndex        =   4
         Top             =   160
         Width           =   330
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "秒后执行关机！"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1200
         TabIndex        =   3
         Top             =   360
         Width           =   1680
      End
      Begin VB.Image Image1 
         Height          =   495
         Left            =   240
         Picture         =   "Shutdn.frx":74E94
         Top             =   240
         Width           =   495
      End
   End
End
Attribute VB_Name = "Shutdn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Option Explicit

Private Sub Form_Load()
i = 5    '关机延迟（秒）
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
i = i - 1
Label3.Caption = i
If i = 0 Then
  Shell "cmd.exe /c shutdown -s -t 0"
  Timer1.Enabled = False
  End
End If
End Sub

Private Sub OKButton_Click()
  Shell "cmd.exe /c shutdown -s -t 0"
End
End Sub

Private Sub CancelButton_Click()
Unload Me
End Sub
