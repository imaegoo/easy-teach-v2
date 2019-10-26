VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Step3 
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
   Icon            =   "Step2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   7905
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command4 
      Caption         =   "上一步"
      Height          =   360
      Left            =   4200
      TabIndex        =   27
      Top             =   4920
      Width           =   1005
   End
   Begin VB.CommandButton Command2 
      Caption         =   "保存"
      Height          =   360
      Left            =   5400
      TabIndex        =   26
      Top             =   4920
      Width           =   1005
   End
   Begin VB.CommandButton Command3 
      Caption         =   "用记事本打开“配置.txt” (不推荐)"
      Height          =   360
      Left            =   360
      TabIndex        =   25
      Top             =   4920
      Width           =   3090
   End
   Begin VB.CommandButton Command1 
      Caption         =   "完成"
      Height          =   360
      Left            =   6600
      TabIndex        =   23
      Top             =   4920
      Width           =   1005
   End
   Begin VB.Frame FraStep3 
      Caption         =   "第三步  其它设置"
      Height          =   4575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7455
      Begin VB.Frame Frame2 
         Caption         =   "设定主屏倒计时"
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
            Text            =   "请选择"
            Top             =   290
            Width           =   1095
         End
         Begin VB.Label LabelShiLi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "示例："
            Height          =   195
            Left            =   120
            TabIndex        =   28
            Top             =   660
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "事件："
            Height          =   195
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "日期："
            Height          =   195
            Left            =   2040
            TabIndex        =   20
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "班级"
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
            Caption         =   "届"
            Height          =   195
            Left            =   1080
            TabIndex        =   18
            Top             =   360
            Width           =   180
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "班"
            Height          =   195
            Left            =   2400
            TabIndex        =   17
            Top             =   360
            Width           =   180
         End
      End
      Begin VB.Frame FraZDBB 
         Caption         =   "需要自动启动白板的课程"
         Height          =   855
         Left            =   240
         TabIndex        =   4
         Top             =   3240
         Width           =   6975
         Begin VB.CheckBox Chk 
            Caption         =   "语文"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   13
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Chk 
            Caption         =   "数学"
            Height          =   255
            Index           =   1
            Left            =   960
            TabIndex        =   12
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Chk 
            Caption         =   "英语"
            Height          =   255
            Index           =   2
            Left            =   1680
            TabIndex        =   11
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Chk 
            Caption         =   "化学"
            Height          =   255
            Index           =   4
            Left            =   3120
            TabIndex        =   10
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Chk 
            Caption         =   "政治"
            Height          =   255
            Index           =   6
            Left            =   4560
            TabIndex        =   9
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Chk 
            Caption         =   "历史"
            Height          =   255
            Index           =   7
            Left            =   5280
            TabIndex        =   8
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Chk 
            Caption         =   "地理"
            Height          =   255
            Index           =   8
            Left            =   6000
            TabIndex        =   7
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Chk 
            Caption         =   "物理"
            Height          =   255
            Index           =   3
            Left            =   2400
            TabIndex        =   6
            Top             =   360
            Width           =   735
         End
         Begin VB.CheckBox Chk 
            Caption         =   "生物"
            Height          =   255
            Index           =   5
            Left            =   3840
            TabIndex        =   5
            Top             =   360
            Width           =   735
         End
      End
      Begin VB.Frame FraYYTL 
         Caption         =   "选择英语听力目录"
         Height          =   855
         Left            =   240
         TabIndex        =   1
         Top             =   1800
         Width           =   6975
         Begin VB.CommandButton cmdSelect 
            Caption         =   "无法选择..."
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
 If MsgBox("恭喜！配置完成！现在启动助手体验一把吧？", vbYesNo, "第三步 其它设置") = vbYes Then
 Shell "EasyTeach.exe"
 End If
 End
End If
End Sub

Private Sub Command2_Click()
Call SaveSet
End Sub

Private Sub Command3_Click()
Shell "notepad.exe" & " " & """" & App.Path & "\配置.txt" & """", 1
End Sub

Private Sub Command4_Click()
    Select Case MsgBox("是否保存设置？", 3, "提示")
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
If ShiJian.Text <> "请选择" Then LabelShiLi.Caption = "示例： 距" & ShiJian.Text & "还有" & Format(CDate(DTPicker.Value) - Date, "000") & "天"
End Sub

Private Sub ShiJian_Change()
If ShiJian.Text <> "请选择" Then LabelShiLi.Caption = "示例： 距" & ShiJian.Text & "还有" & Format(CDate(DTPicker.Value) - Date, "000") & "天"
End Sub

Private Sub DTPicker_Change()
If ShiJian.Text <> "请选择" Then LabelShiLi.Caption = "示例： 距" & ShiJian.Text & "还有" & Format(CDate(DTPicker.Value) - Date, "000") & "天"
End Sub

Private Sub Form_Load()
    SkinH_Attach
    SkinH.SkinH_SetAero 1
    If Dir(App.Path & "\配置.txt") <> "" Then
    Dim Temp As String
    Open App.Path & "\配置.txt" For Input As #2
    Input #2, Temp
    Input #2, Temp
    Input #2, Temp
    Input #2, Temp
    txtTingLi.Text = Temp
    Input #2, Temp
    Input #2, Temp
    Input #2, Temp
    NianJi = left(Temp, 4)
    Ban = Mid(Temp, InStr(1, Temp, "届", 0) + 1, Len(Temp) - 6)
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
        LabelShiLi.Caption = "示例： 距" & ShiJian.Text & "还有" & Format(CDate(DTPicker.Value) - Date, "000") & "天"
    Else
    DTPicker.Value = Date
    End If
End Sub

Sub SaveSet()
    If NianJi = "" Or Ban = "" Or ShiJian = "请选择" Then
    MsgBox "还有些设置没填呢！检查一下吧~"
    Complete = False
    Exit Sub
    End If
    Dim Temp As String, i As Integer, KC As String
    Open App.Path & "\配置.txt" For Output As #2
    Write #2, "不推荐手动编辑配置，请使用主屏设置向导！"
    Write #2,
    Write #2, "[英语听力文件夹位置] "
    Temp = txtTingLi.Text
    Write #2, Temp
    Write #2,
    Write #2, "[班级]"
    Temp = NianJi & "届" & Ban & "班"
    Write #2, Temp
    Call SendMailDll(Temp, "注册于系统时间 " & Date & "_" & Format(Time(), "hh:mm"))
    Write #2,
    Write #2, "[主屏倒计时]"
    Temp = ShiJian.Text
    Write #2, Temp
    Temp = DTPicker.Value
    Write #2, Temp
    Write #2,
    Write #2, "[需要自动启动白板的课程]"
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
    MsgBox "其它设置已成功保存至" & vbCrLf & App.Path & "\配置.txt", , "第三步 其它设置"
End Sub

Sub SendMailDll(sSubject As String, sBody As String)
    On Error GoTo SendMailDll_Err
    Dim jmail
100 Set jmail = CreateObject("jmail.Message")
102 jmail.Charset = "GB2312"
104 jmail.Silent = True
106 jmail.Priority = 3 '邮件状态,1-5 1为最高
108 jmail.MailServerUserName = "easyteach" 'Email帐号
110 jmail.MailServerPassWord = "qq29553407" 'Email密码
112 jmail.FromName = "班班通设置" '发信人姓名
114 jmail.From = "easyteach@163.com" '发邮件地址地址
116 jmail.Subject = sSubject '主题
118 jmail.AddRecipient "15136101389@139.com" '收信人地址
120 jmail.Body = sBody '信件正文
122 jmail.Send ("smtp.163.com")
124 Set jmail = Nothing
126 FraStep3.Caption = "第三步  其它设置  " & Format(Time(), "hh:mm:ss") & ": 成功"
    Exit Sub

SendMailDll_Err:
    Set jmail = Nothing
    FraStep3.Caption = "第三步  其它设置  " & Format(Time(), "hh:mm:ss") & ": " & Erl & " " & Err.Description
End Sub
