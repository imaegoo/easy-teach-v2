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
   StartUpPosition =   2  '��Ļ����
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
        MsgBox "���ǵ�һ��ʹ�ã���������"
        If Dir(App.Path & "\BBTFirst.exe") <> "" Then
            Shell "BBTFirst.exe", 1
        Else
            MsgBox "����㿴���˶Ի���" & vbCrLf & "�Ǳ��ߣ�ɱ�������ɱ�����Ŀ¼���һЩ�ļ�" & vbCrLf & "�뽫��D:\MyTools\����Ӱ���������װһ�飡"
        End If
        End
    End If
    
    If Dir(App.Path & "\�γ̱�.txt") = "" Then
        MsgBox "�γ̱�ʧ���������������ý���"
        Shell "BBTFirst.exe", 1
        End
    End If
    
    If Dir(App.Path & "\����.txt") = "" Then
        MsgBox "�����ļ���ʧ����������������ҳ��"
        Shell "BBTFirst.exe", 1
        End
    End If
    
    FullForm.Show 0
    Unload Me
    
End Sub
