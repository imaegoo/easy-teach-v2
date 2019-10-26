VERSION 5.00
Begin VB.Form BanJiTZ 
   BackColor       =   &H0000FFFF&
   BorderStyle     =   0  'None
   Caption         =   "班级通知"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   FillColor       =   &H80000005&
   BeginProperty Font 
      Name            =   "微软雅黑"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   60
      Top             =   60
   End
   Begin VB.CommandButton CommandOK 
      BackColor       =   &H0000FFFF&
      Caption         =   "确定 (10)"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   1500
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2940
      Width           =   1770
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1650
      TabIndex        =   0
      Top             =   180
      Width           =   1500
   End
End
Attribute VB_Name = "BanJiTZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub CommandOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
If Label1.Caption = "Label1" Then Unload Me
i = 10
End Sub

Private Sub Timer1_Timer()
i = i - 1
CommandOK.Caption = "确定 (" & i & ")"
If i = 0 Then Unload Me
End Sub
