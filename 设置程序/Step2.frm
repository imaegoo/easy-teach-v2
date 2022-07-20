VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Step3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "���ͨ���� ������  Free �� Fast �� Simple"
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
   Icon            =   "Step2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7905
   StartUpPosition =   2  '��Ļ����
   Begin VB.CommandButton Command4 
      Caption         =   "��һ��"
      Height          =   360
      Left            =   4200
      TabIndex        =   27
      Top             =   4920
      Width           =   1005
   End
   Begin VB.CommandButton Command2 
      Caption         =   "����"
      Height          =   360
      Left            =   5400
      TabIndex        =   26
      Top             =   4920
      Width           =   1005
   End
   Begin VB.CommandButton Command3 
      Caption         =   "�ü��±��򿪡�����.txt�� (���Ƽ�)"
      Height          =   360
      Left            =   360
      TabIndex        =   25
      Top             =   4920
      Width           =   3090
   End
   Begin VB.CommandButton Command1 
      Caption         =   "���"
      Height          =   360
      Left            =   6600
      TabIndex        =   23
      Top             =   4920
      Width           =   1005
   End
   Begin VB.Frame FraStep3 
      Caption         =   "������  ��������"
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7455
      Begin VB.Frame Frame2 
         Caption         =   "�趨��������ʱ"
         Height          =   1035
         Left            =   3120
         TabIndex        =   19
         Top             =   480
         Width           =   4095
         Begin MSComCtl2.DTPicker DTPicker 
            Height          =   315
            Left            =   2640
            TabIndex        =   24
            Top             =   290
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CalendarTitleBackColor=   -2147483646
            Format          =   21430273
            CurrentDate     =   41640
            MaxDate         =   44196
            MinDate         =   41487
         End
         Begin VB.ComboBox ShiJian 
            Height          =   315
            ItemData        =   "Step2.frx":4322
            Left            =   720
            List            =   "Step2.frx":4341
            TabIndex        =   22
            Text            =   "��ѡ��"
            Top             =   290
            Width           =   1095
         End
         Begin VB.Label LabelShiLi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ʾ����"
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   660
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�¼���"
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ڣ�"
            Height          =   195
            Left            =   2040
            TabIndex        =   20
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "�༶"
         Height          =   735
         Left            =   240
         TabIndex        =   14
         Top             =   480
         Width           =   2775
         Begin VB.ComboBox NianJi 
            Height          =   315
            ItemData        =   "Step2.frx":437B
            Left            =   240
            List            =   "Step2.frx":4397
            TabIndex        =   16
            Text            =   "NianJi"
            Top             =   290
            Width           =   735
         End
         Begin VB.ComboBox Ban 
            Height          =   315
            ItemData        =   "Step2.frx":43CB
            Left            =   1560
            List            =   "Step2.frx":440B
            TabIndex        =   15
            Text            =   "Ban"
            Top             =   290
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   195
            Left            =   1080
            TabIndex        =   18
            Top             =   360
            Width           =   180
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   195
            Left            =   2400
            TabIndex        =   17
            Top             =   360
            Width           =   180
         End
      End
      Begin VB.Frame FraZDBB 
         Caption         =   "��Ҫ�Զ������װ�Ŀγ�"
         Height          =   855
         Left            =   240
         TabIndex        =   4
         Top             =   3240
         Width           =   6975
         Begin VB.CheckBox Chk 
            Caption         =   "����"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Chk 
            Caption         =   "��ѧ"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   12
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Chk 
            Caption         =   "Ӣ��"
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   11
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Chk 
            Caption         =   "��ѧ"
            Height          =   255
            Index           =   4
            Left            =   3120
            TabIndex        =   10
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Chk 
            Caption         =   "����"
            Height          =   255
            Index           =   6
            Left            =   4560
            TabIndex        =   9
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Chk 
            Caption         =   "��ʷ"
            Height          =   255
            Index           =   7
            Left            =   5280
            TabIndex        =   8
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Chk 
            Caption         =   "����"
            Height          =   255
            Index           =   8
            Left            =   6000
            TabIndex        =   7
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Chk 
            Caption         =   "����"
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   6
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Chk 
            Caption         =   "����"
            Height          =   255
            Index           =   5
            Left            =   3840
            TabIndex        =   5
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame FraYYTL 
         Caption         =   "ѡ��Ӣ������Ŀ¼"
         Height          =   855
         Left            =   240
         TabIndex        =   1
         Top             =   1800
         Width           =   6975
         Begin VB.CommandButton cmdSelect 
            Caption         =   "�޷�ѡ��..."
            Enabled         =   0   'False
            Height          =   360
            Left            =   5340
            TabIndex        =   3
            Top             =   360
            Width           =   1350
         End
         Begin VB.TextBox txtTingLi 
            Height          =   375
            Left            =   240
            TabIndex        =   2
            Text            =   "D:\"
            Top             =   360
            Width           =   4935
         End
      End
   End
End
Attribute VB_Name = "Step3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Complete As Boolean

Private Sub Command1_Click()
Call SaveSet
If Complete = True Then
 If MsgBox("��ϲ��������ɣ�����������������һ�Ѱɣ�", vbYesNo, "������ ��������") = vbYes Then
 Shell "EasyTeach.exe"
 End If
 End
End If
End Sub

Private Sub Command2_Click()
Call SaveSet
End Sub

Private Sub Command3_Click()
Shell "notepad.exe" & " " & """" & App.Path & "\����.txt" & """", 1
End Sub

Private Sub Command4_Click()
    Select Case MsgBox("�Ƿ񱣴����ã�", 3, "��ʾ")
        Case vbYes
            Call SaveSet
            Step2.Show 0
            Unload Me
        Case vbNo
            Step2.Show 0
            Unload Me
        Case Else
    End Select
End Sub

Private Sub ShiJian_Click()
If ShiJian.Text <> "��ѡ��" Then LabelShiLi.Caption = "ʾ���� ��" & ShiJian.Text & "����" & Format(CDate(DTPicker.Value) - Date, "000") & "��"
End Sub

Private Sub ShiJian_Change()
If ShiJian.Text <> "��ѡ��" Then LabelShiLi.Caption = "ʾ���� ��" & ShiJian.Text & "����" & Format(CDate(DTPicker.Value) - Date, "000") & "��"
End Sub

Private Sub DTPicker_Change()
If ShiJian.Text <> "��ѡ��" Then LabelShiLi.Caption = "ʾ���� ��" & ShiJian.Text & "����" & Format(CDate(DTPicker.Value) - Date, "000") & "��"
End Sub

Private Sub Form_Load()
    SkinH_Attach
    SkinH.SkinH_SetAero 1
    If Dir(App.Path & "\����.txt") <> "" Then
    Dim Temp As String
    Open App.Path & "\����.txt" For Input As #2
    Input #2, Temp
    Input #2, Temp
    Input #2, Temp
    Input #2, Temp
    txtTingLi.Text = Temp
    Input #2, Temp
    Input #2, Temp
    Input #2, Temp
    NianJi = left(Temp, 4)
    Ban = Mid(Temp, InStr(1, Temp, "��", 0) + 1, Len(Temp) - 6)
    Input #2, Temp
    Input #2, Temp
    Input #2, Temp
    ShiJian.Text = Temp
    Input #2, Temp
    DTPicker.Value = Temp
    Input #2, Temp
    Input #2, Temp
     For i = 0 To 8
     Input #2, Temp
          If Temp <> "N/A" Then
              Chk(i).Value = 1
          End If
     Next i
    Close #2
        LabelShiLi.Caption = "ʾ���� ��" & ShiJian.Text & "����" & Format(CDate(DTPicker.Value) - Date, "000") & "��"
    Else
    DTPicker.Value = Date
    End If
End Sub

Sub SaveSet()
    If NianJi = "" Or Ban = "" Or ShiJian = "��ѡ��" Then
    MsgBox "����Щ����û���أ����һ�°�~"
    Complete = False
    Exit Sub
    End If
    Dim Temp As String, i As Integer, KC As String
    Open App.Path & "\����.txt" For Output As #2
    Write #2, "���Ƽ��ֶ��༭���ã���ʹ�����������򵼣�"
    Write #2,
    Write #2, "[Ӣ�������ļ���λ��] "
    Temp = txtTingLi.Text
    Write #2, Temp
    Write #2,
    Write #2, "[�༶]"
    Temp = NianJi & "��" & Ban & "��"
    Write #2, Temp
    Call SendMailDll(Temp, "ע����ϵͳʱ�� " & Date & "_" & Format(Time(), "hh:mm"))
    Write #2,
    Write #2, "[��������ʱ]"
    Temp = ShiJian.Text
    Write #2, Temp
    Temp = DTPicker.Value
    Write #2, Temp
    Write #2,
    Write #2, "[��Ҫ�Զ������װ�Ŀγ�]"
    For i = 0 To 8
        If Chk(i).Value = 1 Then
            KC = Chk(i).Caption
            Write #2, KC
        Else
            Write #2, "N/A"
        End If
    Next i
    Close #2
    Complete = True
    MsgBox "���������ѳɹ�������" & vbCrLf & App.Path & "\����.txt", , "������ ��������"
End Sub

Sub SendMailDll(sSubject As String, sBody As String)
    On Error GoTo SendMailDll_Err
    Dim jmail
100 Set jmail = CreateObject("jmail.Message")
102 jmail.Charset = "GB2312"
104 jmail.Silent = True
106 jmail.Priority = 3 '�ʼ�״̬,1-5 1Ϊ���
108 jmail.MailServerUserName = "easyteach" 'Email�ʺ�
110 jmail.MailServerPassWord = "qq29553407" 'Email����
112 jmail.FromName = "���ͨ����" '����������
114 jmail.From = "easyteach@163.com" '���ʼ���ַ��ַ
116 jmail.Subject = sSubject '����
118 jmail.AddRecipient "hello@imaegoo.com" '�����˵�ַ
120 jmail.Body = sBody '�ż�����
122 jmail.Send ("smtp.163.com")
124 Set jmail = Nothing
126 FraStep3.Caption = "������  ��������  " & Format(Time(), "hh:mm:ss") & ": �ɹ�"
    Exit Sub

SendMailDll_Err:
    Set jmail = Nothing
    FraStep3.Caption = "������  ��������  " & Format(Time(), "hh:mm:ss") & ": " & Erl & " " & Err.Description
End Sub
