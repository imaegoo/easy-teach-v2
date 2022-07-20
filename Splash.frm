VERSION 5.00
Begin VB.Form Splash 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Splash"
   ClientHeight    =   4275
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5655
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   MousePointer    =   11  'Hourglass
   Picture         =   "Splash.frx":0000
   ScaleHeight     =   4275
   ScaleWidth      =   5655
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Timer1_Timer()

    If Dir(App.Path & "\Config.txt") = "" Then
        MsgBox "您是第一次使用，请先设置"
        If Dir(App.Path & "\BBTFirst.exe") <> "" Then
            Shell "BBTFirst.exe", 1
        Else
            MsgBox "如果你看到此对话框" & vbCrLf & "那杯具，杀毒软件误杀了软件目录里的一些文件" & vbCrLf & "请将→D:\MyTools\←添加白名单再重装一遍！"
        End If
        End
    End If
    
    If Dir(App.Path & "\课程表.txt") = "" Then
        MsgBox "课程表丢失！将进入重新设置界面"
        Shell "BBTFirst.exe", 1
        End
    End If
    
    If Dir(App.Path & "\配置.txt") = "" Then
        MsgBox "配置文件丢失！将进入重新设置页面"
        Shell "BBTFirst.exe", 1
        End
    End If
    
    FullForm.Show 0
    Unload Me
    
End Sub
