VERSION 5.00
Begin VB.Form NewURL 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�������ַ"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5895
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "NewURL.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5895
   StartUpPosition =   1  '����������
   Begin VB.CommandButton Command4 
      Cancel          =   -1  'True
      Caption         =   "ȡ��"
      Height          =   360
      Left            =   3120
      TabIndex        =   7
      Top             =   1620
      Width           =   990
   End
   Begin VB.CommandButton Command3 
      Caption         =   "ȷ��"
      Default         =   -1  'True
      Height          =   360
      Left            =   1800
      TabIndex        =   6
      Top             =   1620
      Width           =   990
   End
   Begin VB.CommandButton Command2 
      Caption         =   "��Ļ����"
      Height          =   360
      Left            =   4680
      TabIndex        =   5
      Top             =   960
      Width           =   990
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��д����"
      Height          =   360
      Left            =   4680
      TabIndex        =   4
      Top             =   240
      Width           =   990
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   720
      TabIndex        =   3
      Text            =   "http://www."
      Top             =   960
      Width           =   3855
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ַ��"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   1020
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ƣ�"
      Height          =   195
      Left            =   180
      TabIndex        =   0
      Top             =   300
      Width           =   540
   End
End
Attribute VB_Name = "NewURL"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Shell "HandInput.exe", 1
End Sub

Private Sub Command2_Click()
On Error Resume Next
Shell "C:\WINDOWS\system32\osk.exe", 1
End Sub

Private Sub Command3_Click()
If Text1 = "" Then MsgBox "��������վ���ƣ�": Exit Sub
If Text2 = "" Or Text2 = "http://www." Then MsgBox "��������ַ��": Exit Sub
Open App.Path & "\URL.txt" For Append As #1
Write #1, Text1
Write #1, Text2
Close #1
MsgBox "��ӳɹ������������Ч��"
Unload Me
End Sub

Private Sub Command4_Click()
Unload Me
End Sub
