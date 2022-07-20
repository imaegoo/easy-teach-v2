VERSION 5.00
Begin VB.Form Step2 
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
   Icon            =   "Step3.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   368
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   527
   StartUpPosition =   2  '屏幕中心
   Begin VB.CommandButton Command5 
      Caption         =   "启动手写输入"
      Default         =   -1  'True
      Height          =   360
      Left            =   6360
      TabIndex        =   0
      Top             =   0
      Width           =   1290
   End
   Begin VB.CommandButton Command4 
      Caption         =   "上一步"
      Height          =   360
      Left            =   4200
      TabIndex        =   71
      Top             =   4920
      Width           =   1005
   End
   Begin VB.CommandButton Command3 
      Caption         =   "用记事本打开“课程表.txt” (不推荐)"
      Height          =   360
      Left            =   360
      TabIndex        =   70
      Top             =   4920
      Width           =   3090
   End
   Begin VB.CommandButton Command2 
      Caption         =   "保存"
      Height          =   360
      Left            =   5400
      TabIndex        =   72
      Top             =   4920
      Width           =   1005
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下一步"
      Height          =   360
      Left            =   6600
      TabIndex        =   73
      Top             =   4920
      Width           =   1005
   End
   Begin VB.Frame FraKCB 
      Caption         =   "第二步  课程表设置  下拉选择 或键入  2个字以内"
      Height          =   4575
      Left            =   240
      TabIndex        =   75
      Top             =   240
      Width           =   7455
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   0
         ItemData        =   "Step3.frx":4322
         Left            =   720
         List            =   "Step3.frx":436E
         TabIndex        =   1
         Text            =   "N/A"
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   1
         ItemData        =   "Step3.frx":4402
         Left            =   720
         List            =   "Step3.frx":444E
         TabIndex        =   2
         Text            =   "N/A"
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   2
         ItemData        =   "Step3.frx":44E2
         Left            =   720
         List            =   "Step3.frx":452E
         TabIndex        =   3
         Text            =   "N/A"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   3
         ItemData        =   "Step3.frx":45C2
         Left            =   720
         List            =   "Step3.frx":460E
         TabIndex        =   4
         Text            =   "N/A"
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   4
         ItemData        =   "Step3.frx":46A2
         Left            =   720
         List            =   "Step3.frx":46EE
         TabIndex        =   5
         Text            =   "N/A"
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   5
         ItemData        =   "Step3.frx":4782
         Left            =   720
         List            =   "Step3.frx":47CE
         TabIndex        =   6
         Text            =   "N/A"
         Top             =   2520
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   6
         ItemData        =   "Step3.frx":4862
         Left            =   720
         List            =   "Step3.frx":48AE
         TabIndex        =   7
         Text            =   "N/A"
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   7
         ItemData        =   "Step3.frx":4942
         Left            =   720
         List            =   "Step3.frx":498E
         TabIndex        =   8
         Text            =   "N/A"
         Top             =   3240
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   8
         ItemData        =   "Step3.frx":4A22
         Left            =   720
         List            =   "Step3.frx":4A6E
         TabIndex        =   9
         Text            =   "N/A"
         Top             =   3720
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   9
         ItemData        =   "Step3.frx":4B02
         Left            =   720
         List            =   "Step3.frx":4B4E
         TabIndex        =   10
         Text            =   "N/A"
         Top             =   4080
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   10
         ItemData        =   "Step3.frx":4BE2
         Left            =   1680
         List            =   "Step3.frx":4C2E
         TabIndex        =   11
         Text            =   "N/A"
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   11
         ItemData        =   "Step3.frx":4CC2
         Left            =   1680
         List            =   "Step3.frx":4D0E
         TabIndex        =   12
         Text            =   "N/A"
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   12
         ItemData        =   "Step3.frx":4DA2
         Left            =   1680
         List            =   "Step3.frx":4DEE
         TabIndex        =   13
         Text            =   "N/A"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   13
         ItemData        =   "Step3.frx":4E82
         Left            =   1680
         List            =   "Step3.frx":4ECE
         TabIndex        =   14
         Text            =   "N/A"
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   14
         ItemData        =   "Step3.frx":4F62
         Left            =   1680
         List            =   "Step3.frx":4FAE
         TabIndex        =   15
         Text            =   "N/A"
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   15
         ItemData        =   "Step3.frx":5042
         Left            =   1680
         List            =   "Step3.frx":508E
         TabIndex        =   16
         Text            =   "N/A"
         Top             =   2520
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   16
         ItemData        =   "Step3.frx":5122
         Left            =   1680
         List            =   "Step3.frx":516E
         TabIndex        =   17
         Text            =   "N/A"
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   17
         ItemData        =   "Step3.frx":5202
         Left            =   1680
         List            =   "Step3.frx":524E
         TabIndex        =   18
         Text            =   "N/A"
         Top             =   3240
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   18
         ItemData        =   "Step3.frx":52E2
         Left            =   1680
         List            =   "Step3.frx":532E
         TabIndex        =   19
         Text            =   "N/A"
         Top             =   3720
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   19
         ItemData        =   "Step3.frx":53C2
         Left            =   1680
         List            =   "Step3.frx":540E
         TabIndex        =   20
         Text            =   "N/A"
         Top             =   4080
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   20
         ItemData        =   "Step3.frx":54A2
         Left            =   2640
         List            =   "Step3.frx":54EE
         TabIndex        =   21
         Text            =   "N/A"
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   21
         ItemData        =   "Step3.frx":5582
         Left            =   2640
         List            =   "Step3.frx":55CE
         TabIndex        =   22
         Text            =   "N/A"
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   22
         ItemData        =   "Step3.frx":5662
         Left            =   2640
         List            =   "Step3.frx":56AE
         TabIndex        =   23
         Text            =   "N/A"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   23
         ItemData        =   "Step3.frx":5742
         Left            =   2640
         List            =   "Step3.frx":578E
         TabIndex        =   24
         Text            =   "N/A"
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   24
         ItemData        =   "Step3.frx":5822
         Left            =   2640
         List            =   "Step3.frx":586E
         TabIndex        =   25
         Text            =   "N/A"
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   25
         ItemData        =   "Step3.frx":5902
         Left            =   2640
         List            =   "Step3.frx":594E
         TabIndex        =   26
         Text            =   "N/A"
         Top             =   2520
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   26
         ItemData        =   "Step3.frx":59E2
         Left            =   2640
         List            =   "Step3.frx":5A2E
         TabIndex        =   27
         Text            =   "N/A"
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   27
         ItemData        =   "Step3.frx":5AC2
         Left            =   2640
         List            =   "Step3.frx":5B0E
         TabIndex        =   28
         Text            =   "N/A"
         Top             =   3240
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   28
         ItemData        =   "Step3.frx":5BA2
         Left            =   2640
         List            =   "Step3.frx":5BEE
         TabIndex        =   29
         Text            =   "N/A"
         Top             =   3720
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   29
         ItemData        =   "Step3.frx":5C82
         Left            =   2640
         List            =   "Step3.frx":5CCE
         TabIndex        =   30
         Text            =   "N/A"
         Top             =   4080
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   30
         ItemData        =   "Step3.frx":5D62
         Left            =   3600
         List            =   "Step3.frx":5DAE
         TabIndex        =   31
         Text            =   "N/A"
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   31
         ItemData        =   "Step3.frx":5E42
         Left            =   3600
         List            =   "Step3.frx":5E8E
         TabIndex        =   32
         Text            =   "N/A"
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   32
         ItemData        =   "Step3.frx":5F22
         Left            =   3600
         List            =   "Step3.frx":5F6E
         TabIndex        =   33
         Text            =   "N/A"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   33
         ItemData        =   "Step3.frx":6002
         Left            =   3600
         List            =   "Step3.frx":604E
         TabIndex        =   34
         Text            =   "N/A"
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   34
         ItemData        =   "Step3.frx":60E2
         Left            =   3600
         List            =   "Step3.frx":612E
         TabIndex        =   35
         Text            =   "N/A"
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   35
         ItemData        =   "Step3.frx":61C2
         Left            =   3600
         List            =   "Step3.frx":620E
         TabIndex        =   36
         Text            =   "N/A"
         Top             =   2520
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   36
         ItemData        =   "Step3.frx":62A2
         Left            =   3600
         List            =   "Step3.frx":62EE
         TabIndex        =   37
         Text            =   "N/A"
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   37
         ItemData        =   "Step3.frx":6382
         Left            =   3600
         List            =   "Step3.frx":63CE
         TabIndex        =   38
         Text            =   "N/A"
         Top             =   3240
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   38
         ItemData        =   "Step3.frx":6462
         Left            =   3600
         List            =   "Step3.frx":64AE
         TabIndex        =   39
         Text            =   "N/A"
         Top             =   3720
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   39
         ItemData        =   "Step3.frx":6542
         Left            =   3600
         List            =   "Step3.frx":658E
         TabIndex        =   40
         Text            =   "N/A"
         Top             =   4080
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   40
         ItemData        =   "Step3.frx":6622
         Left            =   4560
         List            =   "Step3.frx":666E
         TabIndex        =   41
         Text            =   "N/A"
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   41
         ItemData        =   "Step3.frx":6702
         Left            =   4560
         List            =   "Step3.frx":674E
         TabIndex        =   42
         Text            =   "N/A"
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   42
         ItemData        =   "Step3.frx":67E2
         Left            =   4560
         List            =   "Step3.frx":682E
         TabIndex        =   43
         Text            =   "N/A"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   43
         ItemData        =   "Step3.frx":68C2
         Left            =   4560
         List            =   "Step3.frx":690E
         TabIndex        =   44
         Text            =   "N/A"
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   44
         ItemData        =   "Step3.frx":69A2
         Left            =   4560
         List            =   "Step3.frx":69EE
         TabIndex        =   45
         Text            =   "N/A"
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   45
         ItemData        =   "Step3.frx":6A82
         Left            =   4560
         List            =   "Step3.frx":6ACE
         TabIndex        =   46
         Text            =   "N/A"
         Top             =   2520
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   46
         ItemData        =   "Step3.frx":6B62
         Left            =   4560
         List            =   "Step3.frx":6BAE
         TabIndex        =   47
         Text            =   "N/A"
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   47
         ItemData        =   "Step3.frx":6C42
         Left            =   4560
         List            =   "Step3.frx":6C8E
         TabIndex        =   48
         Text            =   "N/A"
         Top             =   3240
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   48
         ItemData        =   "Step3.frx":6D22
         Left            =   4560
         List            =   "Step3.frx":6D6E
         TabIndex        =   49
         Text            =   "N/A"
         Top             =   3720
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   49
         ItemData        =   "Step3.frx":6E02
         Left            =   4560
         List            =   "Step3.frx":6E4E
         TabIndex        =   50
         Text            =   "N/A"
         Top             =   4080
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   50
         ItemData        =   "Step3.frx":6EE2
         Left            =   5520
         List            =   "Step3.frx":6F2E
         TabIndex        =   51
         Text            =   "N/A"
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   51
         ItemData        =   "Step3.frx":6FC2
         Left            =   5520
         List            =   "Step3.frx":700E
         TabIndex        =   52
         Text            =   "N/A"
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   52
         ItemData        =   "Step3.frx":70A2
         Left            =   5520
         List            =   "Step3.frx":70EE
         TabIndex        =   53
         Text            =   "N/A"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   53
         ItemData        =   "Step3.frx":7182
         Left            =   5520
         List            =   "Step3.frx":71CE
         TabIndex        =   54
         Text            =   "N/A"
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   54
         ItemData        =   "Step3.frx":7262
         Left            =   5520
         List            =   "Step3.frx":72AE
         TabIndex        =   55
         Text            =   "N/A"
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   55
         ItemData        =   "Step3.frx":7342
         Left            =   5520
         List            =   "Step3.frx":738E
         TabIndex        =   56
         Text            =   "N/A"
         Top             =   2520
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   56
         ItemData        =   "Step3.frx":7422
         Left            =   5520
         List            =   "Step3.frx":746E
         TabIndex        =   57
         Text            =   "N/A"
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   57
         ItemData        =   "Step3.frx":7502
         Left            =   5520
         List            =   "Step3.frx":754E
         TabIndex        =   58
         Text            =   "N/A"
         Top             =   3240
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   58
         ItemData        =   "Step3.frx":75E2
         Left            =   5520
         List            =   "Step3.frx":762E
         TabIndex        =   59
         Text            =   "N/A"
         Top             =   3720
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   59
         ItemData        =   "Step3.frx":76C2
         Left            =   5520
         List            =   "Step3.frx":770E
         TabIndex        =   60
         Text            =   "N/A"
         Top             =   4080
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   60
         ItemData        =   "Step3.frx":77A2
         Left            =   6480
         List            =   "Step3.frx":77EE
         TabIndex        =   61
         Text            =   "N/A"
         Top             =   600
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   61
         ItemData        =   "Step3.frx":7882
         Left            =   6480
         List            =   "Step3.frx":78CE
         TabIndex        =   62
         Text            =   "N/A"
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   62
         ItemData        =   "Step3.frx":7962
         Left            =   6480
         List            =   "Step3.frx":79AE
         TabIndex        =   63
         Text            =   "N/A"
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   63
         ItemData        =   "Step3.frx":7A42
         Left            =   6480
         List            =   "Step3.frx":7A8E
         TabIndex        =   64
         Text            =   "N/A"
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   64
         ItemData        =   "Step3.frx":7B22
         Left            =   6480
         List            =   "Step3.frx":7B6E
         TabIndex        =   65
         Text            =   "N/A"
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   65
         ItemData        =   "Step3.frx":7C02
         Left            =   6480
         List            =   "Step3.frx":7C4E
         TabIndex        =   66
         Text            =   "N/A"
         Top             =   2520
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   66
         ItemData        =   "Step3.frx":7CE2
         Left            =   6480
         List            =   "Step3.frx":7D2E
         TabIndex        =   67
         Text            =   "N/A"
         Top             =   2880
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   67
         ItemData        =   "Step3.frx":7DC2
         Left            =   6480
         List            =   "Step3.frx":7E0E
         TabIndex        =   68
         Text            =   "N/A"
         Top             =   3240
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   68
         ItemData        =   "Step3.frx":7EA2
         Left            =   6480
         List            =   "Step3.frx":7EEE
         TabIndex        =   69
         Text            =   "N/A"
         Top             =   3720
         Width           =   855
      End
      Begin VB.ComboBox Com 
         Height          =   315
         Index           =   69
         ItemData        =   "Step3.frx":7F82
         Left            =   6480
         List            =   "Step3.frx":7FCE
         TabIndex        =   74
         Text            =   "N/A"
         Top             =   4080
         Width           =   855
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "周一"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   720
         TabIndex        =   92
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "周二"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   1680
         TabIndex        =   91
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "周三"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   2640
         TabIndex        =   90
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "周四"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   3600
         TabIndex        =   89
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "周五"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   4560
         TabIndex        =   88
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "周六"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   5520
         TabIndex        =   87
         Top             =   240
         Width           =   480
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "周日"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   6480
         TabIndex        =   86
         Top             =   240
         Width           =   480
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "AM1"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   0
         Left            =   30
         TabIndex        =   85
         Top             =   600
         Width           =   540
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   420
         TabIndex        =   84
         Top             =   960
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   420
         TabIndex        =   83
         Top             =   1320
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   420
         TabIndex        =   82
         Top             =   1680
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   420
         TabIndex        =   81
         Top             =   2040
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PM1"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   45
         TabIndex        =   80
         Top             =   2520
         Width           =   525
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   6
         Left            =   420
         TabIndex        =   79
         Top             =   2880
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   7
         Left            =   420
         TabIndex        =   78
         Top             =   3240
         Width           =   135
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "NT1"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   8
         Left            =   105
         TabIndex        =   77
         Top             =   3720
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "微软雅黑"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   9
         Left            =   420
         TabIndex        =   76
         Top             =   4080
         Width           =   135
      End
   End
End
Attribute VB_Name = "Step2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
SkinH_Attach
SkinH.SkinH_SetAero 1
    Dim KC As String, Temp As String
If Dir(App.Path & "\课程表.txt") <> "" Then
    Open App.Path & "\课程表.txt" For Input As #1
    Input #1, Temp
    Input #1, Temp
    Input #1, Temp
    Input #1, Temp
    Input #1, Temp
    Input #1, Temp

    For Week = 0 To 6                    '周一 ～ 周日  对应数值  0～6
        Input #1, Temp
        Input #1, KC
        Com(Week * 10 + 0).Text = KC
        Input #1, KC
        Com(Week * 10 + 1).Text = KC
        Input #1, KC
        Com(Week * 10 + 2).Text = KC
        Input #1, KC
        Com(Week * 10 + 3).Text = KC
        Input #1, KC
        Input #1, Temp
        Com(Week * 10 + 4).Text = KC
        Input #1, KC
        Com(Week * 10 + 5).Text = KC
        Input #1, KC
        Com(Week * 10 + 6).Text = KC
        Input #1, KC
        Input #1, Temp
        Com(Week * 10 + 7).Text = KC
        Input #1, KC
        Com(Week * 10 + 8).Text = KC
        Input #1, KC
        Com(Week * 10 + 9).Text = KC
        Input #1, Temp
    Next Week

    Close #1
End If

End Sub

Private Sub Command1_Click()
    Select Case MsgBox("是否保存设置？", 3, "提示")
        Case vbYes
            Call SaveKCB
            Step3.Show 0
            Unload Me
        Case vbNo
            Step3.Show 0
            Unload Me
        Case Else
    End Select
End Sub

Private Sub Command2_Click()
Call SaveKCB
End Sub

Private Sub Command3_Click()
Shell "notepad.exe" & " " & """" & App.Path & "\课程表.txt" & """", 1
End Sub

Private Sub Command4_Click()

    Select Case MsgBox("是否保存设置？", 3, "提示")

        Case vbYes
            Call SaveKCB
            Step1.Show 0
            Unload Me

        Case vbNo
            Step1.Show 0
            Unload Me

        Case Else
    End Select

End Sub

Private Sub Command5_Click()
Shell "HandInput.exe", 1
End Sub

Sub SaveKCB()

    If MsgBox("此操作将会覆盖以前保存的课程表，确定继续？", 1, "警告") = vbOK Then
        Dim WeekCHS As String, KC As String
        Open App.Path & "\课程表.txt" For Output As #1
        Write #1, "注意：由于班班通助手使用读取指定行的方式来获取课程表，编辑此文件时请不要弄乱原本的行数！"
        Write #1, "　　　否则在助手中就会显示错位，甚至崩溃！推荐使用设置向导来设置课程表！！"
        Write #1,
        Write #1, "[课程表标题]"
        Write #1, "智能课程表"
        Write #1,

        For Week = 0 To 6                    '周一 ～ 周日  对应数值  0～6

            Select Case Week

                Case 0
                    WeekCHS = "周一"

                Case 1
                    WeekCHS = "周二"

                Case 2
                    WeekCHS = "周三"

                Case 3
                    WeekCHS = "周四"

                Case 4
                    WeekCHS = "周五"

                Case 5
                    WeekCHS = "周六"

                Case 6
                    WeekCHS = "周日"
            End Select

            Write #1, "[" & WeekCHS & "]"
            KC = Com(Week * 10 + 0).Text
            Write #1, KC
            KC = Com(Week * 10 + 1).Text
            Write #1, KC
            KC = Com(Week * 10 + 2).Text
            Write #1, KC
            KC = Com(Week * 10 + 3).Text
            Write #1, KC
            KC = Com(Week * 10 + 4).Text
            Write #1, KC
            Write #1,
            KC = Com(Week * 10 + 5).Text
            Write #1, KC
            KC = Com(Week * 10 + 6).Text
            Write #1, KC
            KC = Com(Week * 10 + 7).Text
            Write #1, KC
            Write #1,
            KC = Com(Week * 10 + 8).Text
            Write #1, KC
            KC = Com(Week * 10 + 9).Text
            Write #1, KC
            Write #1,
        Next Week

        Close #1
        MsgBox "课程表已成功保存至" & vbCrLf & App.Path & "\课程表.txt", , "第二步 课程表"
    End If

End Sub
