VERSION 5.00
Begin VB.Form Password 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����������"
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
   StartUpPosition =   1  '����������
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
      Caption         =   "С����"
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
      Caption         =   "����룺000   ʱ�䣺00:00"
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
      Caption         =   "���붯̬�����Խ������ܡ���"
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
'�ڼ�������ʱ�ɹ�ʹ�õı�����
'It     ����룬ȡֵ100~999������
'It1     ������λ��ȡֵ1~9������
'It2     �����ʮλ��ȡֵ0~9������
'It3     ������λ��ȡֵ0~9������
'Hr     ϵͳСʱ����0~23������
'Hr1    Сʱ��ʮλ��0~2������
'Hr2    Сʱ����λ��0~9������
'Mn     ��������0~59������
'Mn1    ������ʮλ��0~5������
'Mn2    ��������λ��0~9������

Function Pwd1True(Pwd As Integer) As Boolean
Dim TruePwd As Integer
TruePwd = (It1 + Hr1) Mod 10                                      '��һλ�㷨
'�ҵ�����Ϊ��������λ+Сʱ��ʮλ�����ȡ��λ������10ȡ�಻����ȡ��λô��
'��������ڶ�λʲô�����ư�
If Pwd = TruePwd Then
Pwd1True = True
Else
Pwd1True = False
End If
End Function

Function Pwd2True(Pwd As Integer) As Boolean
Dim TruePwd As Integer
TruePwd = Abs((It1 - Hr1) Mod 10)                                 '�ڶ�λ�㷨
If Pwd = TruePwd Then
Pwd2True = True
Else
Pwd2True = False
End If
End Function

Function Pwd3True(Pwd As Integer) As Boolean
Dim TruePwd As Integer
TruePwd = (It3 + Mn2) Mod 10                                      '����λ�㷨
If Pwd = TruePwd Then
Pwd3True = True
Else
Pwd3True = False
End If
End Function

Function Pwd4True(Pwd As Integer) As Boolean
Dim TruePwd As Integer
TruePwd = Abs((It3 - Mn2) Mod 10)                                 '����λ�㷨
If Pwd = TruePwd Then
Pwd4True = True
Else
Pwd4True = False
End If
End Function

Private Sub Form_Load()
SkinH_Attach
SkinH.SkinH_SetAero 1
Randomize   '��ʼ�������
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
LabelInt.Caption = "����룺" & It & "   ʱ�䣺" & Format(Hr, "00") & ":" & Format(Mn, "00")
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
      Label1.Caption = "����������ȷ"
      '����������������ȷʱ��ɶ������
      Shell "C:\PROGRA~1\VCOMTO~1\�����ӿ�V1.01.exe"
      If MsgBox("����������Ӧ�ñ��ζԵ��Ե����и��ģ�", vbYesNo, "д�����ѿ���") = vbYes Then
      Shell "cmd.exe /c shutdown -r -t 0"
      End If
      End
      '����������������ȷʱ��ɶ������
      Else
      Label1.Caption = "�����������"
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
