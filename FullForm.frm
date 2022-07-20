VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Begin VB.Form FullForm 
   BackColor       =   &H00946E50&
   BorderStyle     =   0  'None
   Caption         =   "���ͨ���� II"
   ClientHeight    =   11070
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15360
   Icon            =   "FullForm.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   738
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   Begin VB.Timer TimerUSB 
      Enabled         =   0   'False
      Interval        =   6000
      Left            =   3240
      Top             =   2280
   End
   Begin VB.PictureBox PicturePM 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   0
      Picture         =   "FullForm.frx":4322
      ScaleHeight     =   315
      ScaleWidth      =   15360
      TabIndex        =   78
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   15360
      Begin VB.Label LabelPMZS2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "�ض���Ⱦ"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   14280
         TabIndex        =   80
         Top             =   0
         Width           =   960
      End
      Begin VB.Label LabelPMZS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��������:��N/A"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   60
         TabIndex        =   79
         Top             =   0
         Width           =   1500
      End
   End
   Begin SysInfoLib.SysInfo SysInfo1 
      Left            =   120
      Top             =   240
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.CommandButton CommandJKKG 
      BackColor       =   &H005EDFBF&
      Caption         =   "�ӿڿ���"
      Height          =   375
      Left            =   10560
      MaskColor       =   &H005EDFBF&
      Style           =   1  'Graphical
      TabIndex        =   74
      Top             =   10320
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton CommandNB 
      BackColor       =   &H005EDFBF&
      Caption         =   "��ѧ������"
      Height          =   375
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   72
      Top             =   10320
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Timer TimerLoad 
      Interval        =   10
      Left            =   120
      Top             =   840
   End
   Begin VB.CommandButton CommandQvod 
      BackColor       =   &H005EDFBF&
      Caption         =   "���λ���"
      Height          =   375
      Left            =   8160
      Style           =   1  'Graphical
      TabIndex        =   64
      Top             =   10320
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox PictureMusicWait 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   8760
      Picture         =   "FullForm.frx":16364
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   62
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   2250
      Begin VB.Timer TimerMusicWait 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   120
         Top             =   120
      End
   End
   Begin VB.PictureBox PictureQvodWait 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   6120
      Picture         =   "FullForm.frx":26C7E
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   61
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   2250
      Begin VB.Timer TimerQvodWait 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   120
         Top             =   120
      End
   End
   Begin VB.PictureBox PictureBaiduWait 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   8760
      Picture         =   "FullForm.frx":37598
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   60
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   2250
      Begin VB.Timer TimerBaiduWait 
         Enabled         =   0   'False
         Interval        =   10000
         Left            =   120
         Top             =   120
      End
   End
   Begin VB.PictureBox PicturePPTWait 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   6120
      Picture         =   "FullForm.frx":47EB2
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   2250
      Begin VB.Timer TimerPPTWait 
         Enabled         =   0   'False
         Interval        =   20000
         Left            =   120
         Top             =   120
      End
   End
   Begin VB.CommandButton CommandD 
      BackColor       =   &H005EDFBF&
      Caption         =   "��Ӳ��"
      Default         =   -1  'True
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   10320
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton CommandAutoShut 
      BackColor       =   &H005EDFBF&
      Caption         =   "�Զ��ػ�"
      Height          =   375
      Left            =   6960
      MaskColor       =   &H005EDFBF&
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   10320
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.PictureBox PictureAutoBoard 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   600
      Picture         =   "FullForm.frx":587CC
      ScaleHeight     =   2250
      ScaleWidth      =   4890
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   5280
      Visible         =   0   'False
      Width           =   4890
      Begin VB.Timer TimerAutoBoard 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   120
         Top             =   120
      End
      Begin VB.Label LabelText6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ȡ��"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   570
         Left            =   1560
         TabIndex        =   21
         Top             =   1200
         Width           =   1740
      End
      Begin VB.Label LabelText5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10���򿪰װ�"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   570
         Left            =   900
         TabIndex        =   20
         Top             =   480
         Width           =   3120
      End
   End
   Begin VB.Timer TimerMinute 
      Interval        =   60000
      Left            =   1800
      Top             =   240
   End
   Begin VB.PictureBox PictureBaidu 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   8760
      Picture         =   "FullForm.frx":7C646
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2250
   End
   Begin VB.PictureBox PictureBaiduNo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   8760
      Picture         =   "FullForm.frx":8CF60
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2250
   End
   Begin VB.PictureBox PictureUSBYes 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   3240
      Picture         =   "FullForm.frx":9D87A
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   2250
      Begin VB.Label LabelUSBDrive 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F:\"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   1320
         TabIndex        =   75
         Top             =   1335
         Width           =   315
      End
      Begin VB.Label LabelRemove 
         Alignment       =   2  'Center
         BackColor       =   &H002E28D1&
         Caption         =   " �� �����Ƴ� �� "
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   330
         TabIndex        =   58
         Top             =   0
         Width           =   1620
      End
   End
   Begin VB.Timer TimerSecond 
      Interval        =   1000
      Left            =   1320
      Top             =   240
   End
   Begin VB.PictureBox PictureUnload 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   9960
      Picture         =   "FullForm.frx":AE194
      ScaleHeight     =   1050
      ScaleWidth      =   1050
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1050
   End
   Begin VB.PictureBox PictureMin 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   8760
      Picture         =   "FullForm.frx":B1BCE
      ScaleHeight     =   1050
      ScaleWidth      =   1050
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1050
   End
   Begin VB.PictureBox PictureRestart 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   8760
      Picture         =   "FullForm.frx":B5608
      ScaleHeight     =   1050
      ScaleWidth      =   1050
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1050
   End
   Begin VB.PictureBox PictureShut 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   9960
      Picture         =   "FullForm.frx":B9042
      ScaleHeight     =   1050
      ScaleWidth      =   1050
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1050
   End
   Begin VB.PictureBox PictureKeyboard 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   7320
      Picture         =   "FullForm.frx":BCA7C
      ScaleHeight     =   1050
      ScaleWidth      =   1050
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1050
   End
   Begin VB.PictureBox PictureHand 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   6120
      Picture         =   "FullForm.frx":C04B6
      ScaleHeight     =   1050
      ScaleWidth      =   1050
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7800
      Width           =   1050
   End
   Begin VB.PictureBox PictureVcom 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   6120
      Picture         =   "FullForm.frx":C3EF0
      ScaleHeight     =   1050
      ScaleWidth      =   1050
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1050
   End
   Begin VB.PictureBox PictureMore 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1050
      Left            =   7320
      Picture         =   "FullForm.frx":C792A
      ScaleHeight     =   1050
      ScaleWidth      =   1050
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   9000
      Width           =   1050
      Begin VB.Timer TimerStart 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   120
         Top             =   120
      End
   End
   Begin VB.PictureBox PictureMusic 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   8760
      Picture         =   "FullForm.frx":CB364
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   5280
      Width           =   2250
   End
   Begin VB.PictureBox PictureQvod 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   6120
      Picture         =   "FullForm.frx":DBC7E
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   5280
      Width           =   2250
   End
   Begin VB.PictureBox PicturePPT 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   6120
      Picture         =   "FullForm.frx":EC598
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2250
   End
   Begin VB.PictureBox PictureBoard 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   600
      Picture         =   "FullForm.frx":FCEB2
      ScaleHeight     =   2250
      ScaleWidth      =   4890
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5280
      Width           =   4890
      Begin VB.Timer TimerBoard 
         Enabled         =   0   'False
         Interval        =   500
         Left            =   120
         Top             =   120
      End
      Begin VB.Label LabelKC 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   420
         Left            =   360
         TabIndex        =   71
         Top             =   780
         Width           =   660
      End
      Begin VB.Label LabelZJS 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ͨ����ȫ�Զ��򿪰װ�^_^"
         BeginProperty Font 
            Name            =   "΢���ź�"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   360
         TabIndex        =   63
         Top             =   1680
         Width           =   2505
      End
   End
   Begin VB.PictureBox PictureUSBNo 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   3240
      Picture         =   "FullForm.frx":120D2C
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2250
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   780
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "��ѡ����Ҫ��ӳ�Ļõ�Ƭ"
      Flags           =   12
      InitDir         =   "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}"
   End
   Begin VB.PictureBox PictureComputerWait 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   600
      Picture         =   "FullForm.frx":131646
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   57
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   2250
      Begin VB.Timer TimerComputerWait 
         Enabled         =   0   'False
         Interval        =   5000
         Left            =   120
         Top             =   120
      End
   End
   Begin VB.PictureBox PictureComputer 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2250
      Left            =   600
      Picture         =   "FullForm.frx":141F60
      ScaleHeight     =   2250
      ScaleWidth      =   2250
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2760
      Width           =   2250
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser2 
      Height          =   1695
      Left            =   900
      TabIndex        =   73
      Top             =   3060
      Width           =   1695
      ExtentX         =   2990
      ExtentY         =   2990
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.ComboBox ComboURL 
      BackColor       =   &H007015EA&
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      ItemData        =   "FullForm.frx":15287A
      Left            =   9120
      List            =   "FullForm.frx":152890
      TabIndex        =   77
      Text            =   "Choose"
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00:00"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   44.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1170
      Index           =   1
      Left            =   11700
      TabIndex        =   81
      Top             =   480
      Width           =   2310
   End
   Begin VB.Label lblMail 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����/�������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   11205
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   76
      Top             =   10275
      Width           =   1935
   End
   Begin VB.Label LabelDate 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Date: N/A"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   30
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   780
      Left            =   11850
      TabIndex        =   70
      Top             =   1440
      Width           =   2850
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ڻ�ȡ���� ����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   30
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   780
      Index           =   2
      Left            =   600
      TabIndex        =   69
      Top             =   1440
      Width           =   4770
   End
   Begin VB.Label LabelTips 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "���ͨ���� v0.0 N/A�桡����ʱ��: "
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   600
      TabIndex        =   68
      Top             =   10275
      Width           =   4830
   End
   Begin VB.Label LabelXQ 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   13440
      MousePointer    =   14  'Arrow and Question
      TabIndex        =   67
      Top             =   10275
      Width           =   1200
   End
   Begin VB.Label lblStart 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼ Start"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   44.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1170
      Index           =   1
      Left            =   600
      TabIndex        =   66
      Top             =   465
      Width           =   4035
   End
   Begin VB.Label LabelXingQi 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "֪ͨ����+״̬"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   12660
      TabIndex        =   65
      Top             =   2160
      Width           =   2025
   End
   Begin VB.Label LabelGaoKao 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   39
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   3000
      TabIndex        =   56
      Top             =   8880
      Width           =   1335
   End
   Begin VB.Label LabelText3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��߿�����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1320
      TabIndex        =   55
      Top             =   9120
      Width           =   1575
   End
   Begin VB.Label LabelText1 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��N/A����"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   1335
      TabIndex        =   54
      Top             =   8280
      Width           =   1560
   End
   Begin VB.Label LabelText4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   4440
      TabIndex        =   53
      Top             =   9120
      Width           =   315
   End
   Begin VB.Label LabelKaoShi 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   39
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   915
      Left            =   3000
      TabIndex        =   52
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label LabelText2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   4440
      TabIndex        =   51
      Top             =   8280
      Width           =   315
   End
   Begin VB.Shape ShapeControl 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00804000&
      FillStyle       =   0  'Solid
      Height          =   2250
      Left            =   600
      Top             =   7800
      Width           =   4905
   End
   Begin VB.Label LabelKCTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "__�� ���տγ�"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   11880
      TabIndex        =   50
      Top             =   2880
      Width           =   1785
   End
   Begin VB.Label LabelAM 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "========����========"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   11940
      TabIndex        =   49
      Top             =   3480
      Width           =   2370
   End
   Begin VB.Label LabelAM1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   11880
      TabIndex        =   48
      Top             =   3720
      Width           =   1080
   End
   Begin VB.Label LabelAM2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   11880
      TabIndex        =   47
      Top             =   4320
      Width           =   1080
   End
   Begin VB.Label LabelAM3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   11880
      TabIndex        =   46
      Top             =   4920
      Width           =   1080
   End
   Begin VB.Label LabelAM4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   11880
      TabIndex        =   45
      Top             =   5520
      Width           =   1080
   End
   Begin VB.Label LabelAM5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   11880
      TabIndex        =   44
      Top             =   6120
      Width           =   1080
   End
   Begin VB.Label LabelPM 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "=========����========"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   11820
      TabIndex        =   43
      Top             =   6600
      Width           =   2490
   End
   Begin VB.Label LabelPM1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   11880
      TabIndex        =   42
      Top             =   6840
      Width           =   1080
   End
   Begin VB.Label LabelPM2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   11880
      TabIndex        =   41
      Top             =   7440
      Width           =   1080
   End
   Begin VB.Label LabelPM3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   11880
      TabIndex        =   40
      Top             =   8040
      Width           =   1080
   End
   Begin VB.Label LabelToday 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   12180
      TabIndex        =   39
      Top             =   3360
      Width           =   450
   End
   Begin VB.Label LabelNight 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "========����ϰ======="
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   11820
      TabIndex        =   38
      Top             =   8520
      Width           =   2490
   End
   Begin VB.Label LabelNT1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   11880
      TabIndex        =   37
      Top             =   8760
      Width           =   1080
   End
   Begin VB.Label LabelNT2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   525
      Left            =   11880
      TabIndex        =   36
      Top             =   9360
      Width           =   1080
   End
   Begin VB.Label LabelTomorrow 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDECBF&
      Height          =   210
      Left            =   13620
      TabIndex        =   35
      Top             =   3360
      Width           =   450
   End
   Begin VB.Label LabelAMT1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDECBF&
      Height          =   525
      Left            =   13260
      TabIndex        =   34
      Top             =   3720
      Width           =   1080
   End
   Begin VB.Label LabelAMT2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDECBF&
      Height          =   525
      Left            =   13260
      TabIndex        =   33
      Top             =   4320
      Width           =   1080
   End
   Begin VB.Label LabelAMT3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDECBF&
      Height          =   525
      Left            =   13260
      TabIndex        =   32
      Top             =   4920
      Width           =   1080
   End
   Begin VB.Label LabelAMT4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDECBF&
      Height          =   525
      Left            =   13260
      TabIndex        =   31
      Top             =   5520
      Width           =   1080
   End
   Begin VB.Label LabelAMT5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDECBF&
      Height          =   525
      Left            =   13260
      TabIndex        =   30
      Top             =   6120
      Width           =   1080
   End
   Begin VB.Label LabelPMT1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDECBF&
      Height          =   525
      Left            =   13260
      TabIndex        =   29
      Top             =   6840
      Width           =   1080
   End
   Begin VB.Label LabelPMT2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDECBF&
      Height          =   525
      Left            =   13260
      TabIndex        =   28
      Top             =   7440
      Width           =   1080
   End
   Begin VB.Label LabelPMT3 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDECBF&
      Height          =   525
      Left            =   13260
      TabIndex        =   27
      Top             =   8040
      Width           =   1080
   End
   Begin VB.Label LabelNTT1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDECBF&
      Height          =   525
      Left            =   13260
      TabIndex        =   26
      Top             =   8760
      Width           =   1080
   End
   Begin VB.Label LabelNTT2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "����"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FDECBF&
      Height          =   525
      Left            =   13260
      TabIndex        =   25
      Top             =   9360
      Width           =   1080
   End
   Begin VB.Label LabelWeek 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��__"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   390
      Left            =   13860
      TabIndex        =   24
      Top             =   2880
      Width           =   570
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   876
      X2              =   876
      Y1              =   264
      Y2              =   440
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   876
      X2              =   876
      Y1              =   456
      Y2              =   560
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   876
      X2              =   876
      Y1              =   584
      Y2              =   656
   End
   Begin VB.Label LabelAbout 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "����Ԥ��"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   405
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   1200
   End
   Begin VB.Label lblTime 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "΢���ź�"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Index           =   0
      Left            =   14160
      TabIndex        =   0
      Top             =   615
      Width           =   510
   End
   Begin VB.Shape ShapeKCB 
      BackStyle       =   1  'Opaque
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00D9A108&
      FillStyle       =   0  'Solid
      Height          =   7305
      Left            =   11640
      Top             =   2760
      Width           =   3015
   End
End
Attribute VB_Name = "FullForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'==����API===================================================================
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'˯��
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
'���ģ��
Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'���ģ�����
Private Declare Function GetPixel Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long) As Long
'��ȡ��Ļĳ����ɫ
Private Declare Function GetWindowDC Lib "user32" (ByVal hWnd As Long) As Long
'��ȡ��Ļĳ����ɫ���

Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
'���ģ�����

Private Declare Sub ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long)
Private Const SW_SHOWNORMAL = 1
'����ҳ

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
    ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, _
    ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Type POINTAPI    '�����û��Զ��������
    X As Long    '����X����������
    Y As Long    '����Y����������
End Type
'��Ļ�������ɫ���

'==����ȫ�ֱ���==============================================================
Dim AM1, AM2, AM3, AM4, AM5, PM1, PM2, PM3, NT1, NT2 As String
Dim AMT1, AMT2, AMT3, AMT4, AMT5, PMT1, PMT2, PMT3, NTT1, NTT2 As String
'�γ̱��ַ�����, AM����, PM����, NT����ϰ, ��T��ʾ�ڶ����
Dim KCTitle, tm As String, daojs As Integer, USBa As Integer
'����Ϊ: �γ̱���Ŀ, ��ǰʱ��, �װ忪���Զ���������ʱ
Dim cor As Long, iBoard As Single, hang As Integer
'����Ϊ: ��ȡ����Ļ��ɫ, �����װ�����ʱ, ��ȡ����ʱ����������
Dim a As Single, Weatheri As Integer
'Tips����ʱ, ����10����һ���µļ���
Dim ErrDesc As String, TqBj As Integer
'��������
Dim BanJi As String, ZhuPingDJS As String, ZhuPingDJSDate As Date
Dim Environment1, Environment2 As Long
Dim BB1, BB2, BB3, BB4, BB5, BB6, BB7, BB8, BB9, BBAll As String
Dim TVersion, TVersionText As String
Dim GongGaoComplete As Boolean, USBIsOn As Boolean
Dim URLTitle(20) As String, URL(20) As String
'�ã�����

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

'==����ʱִ��================================================================
Private Sub TimerLoad_Timer()
    On Error Resume Next
    TVersion = "2.7.0"
    TVersionText = "Ver.Final"
    daojs = 11
    kjtm = Time()
If Dir(App.Path & "\��������\") <> "" Then
    Me.Picture = LoadPicture(App.Path & "\��������\Ĭ�ϱ���.jpg")
End If
    Call LoadPeizhi
    Call LoadConfig
    Call LoadKCB
    Call LoadWebBrowser
    Call LoadUI
    Call LoadURL
    TimerLoad.Enabled = False
End Sub

Sub LoadPeizhi()
        '<EhHeader>
        On Error GoTo LoadPeizhi_Err
        '</EhHeader>
        Dim TempStr As String
100     Open App.Path & "\����.txt" For Input As #1
102     Input #1, TempStr
104     Input #1, TempStr
106     Input #1, TempStr
108     Input #1, TempStr
110     Input #1, TempStr
112     Input #1, TempStr
114     Input #1, BanJi                                                     '�༶
116     Input #1, TempStr
118     Input #1, TempStr
120     Input #1, ZhuPingDJS
122     Input #1, TempStr
124     ZhuPingDJSDate = CDate(TempStr)                                   '����ʱ
126     Input #1, TempStr
128     Input #1, TempStr
130     Input #1, BB1
132     Input #1, BB2
134     Input #1, BB3
136     Input #1, BB4
138     Input #1, BB5
140     Input #1, BB6
142     Input #1, BB7
144     Input #1, BB8
146     Input #1, BB9
156     Close #1
158     If BB1 <> "N/A" Then
160         BBAll = BBAll & " " & BB1
        End If
162     If BB2 <> "N/A" Then
164         BBAll = BBAll & " " & BB2
        End If
166     If BB3 <> "N/A" Then
168         BBAll = BBAll & " " & BB3
        End If
170     If BB4 <> "N/A" Then
172         BBAll = BBAll & " " & BB4
        End If
174     If BB5 <> "N/A" Then
176         BBAll = BBAll & " " & BB5
        End If
178     If BB6 <> "N/A" Then
180         BBAll = BBAll & " " & BB6
        End If
182     If BB7 <> "N/A" Then
184         BBAll = BBAll & " " & BB7
        End If
186     If BB8 <> "N/A" Then
188         BBAll = BBAll & " " & BB8
        End If
190     If BB9 <> "N/A" Then
192         BBAll = BBAll & " " & BB9
        End If
        If BBAll <> "" Then
194     BBAll = Right(BBAll, Len(BBAll) - 1)
196     LabelKC = BBAll
        End If
        '<EhFooter>
        Exit Sub

LoadPeizhi_Err:
        MsgBox Err.Description & vbCrLf & _
               "����.txt ���س��������¼��Щ��Ϣ�����������ɻ�QQ29553407" & _
               "�����" & Erl, _
               vbExclamation + vbOKOnly, "�ܱ�Ǹ����������"
        Call SendMailDll(BanJi & "����", Err.Description & " in LoadPeizhi at " & Erl)
        Resume Next
        '</EhFooter>
End Sub

Sub LoadConfig()
        '<EhHeader>
        On Error GoTo LoadConfig_Err
        '</EhHeader>
100     Open App.Path & "\Config.txt" For Input As #1
102     Input #1, TempStr
104     Input #1, Environment1
106     Input #1, Environment2
108     Close #1
        '<EhFooter>
        Exit Sub

LoadConfig_Err:
        MsgBox Err.Description & vbCrLf & _
               "Config.txt ���س��������¼��Щ��Ϣ�����������ɻ�QQ29553407" & _
               "�����" & Erl, _
               vbExclamation + vbOKOnly, "�ܱ�Ǹ����������"
        Call SendMailDll(BanJi & "����", Err.Description & " in LoadConfig at " & Erl)
        Resume Next
        '</EhFooter>
End Sub

Sub LoadKCB()
        '<EhHeader>
        On Error GoTo LoadKCB_Err
        '</EhHeader>
        Dim week As Integer
100     week = Weekday(Now, vbMonday)

102     Select Case (week)

            Case 1
104             LabelWeek.Caption = "��һ"

106         Case 2
108             LabelWeek.Caption = "�ܶ�"

110         Case 3
112             LabelWeek.Caption = "����"

114         Case 4
116             LabelWeek.Caption = "����"

118         Case 5
120             LabelWeek.Caption = "����"

122         Case 6
124             LabelWeek.Caption = "����"

126         Case 7
128             LabelWeek.Caption = "����"
        End Select
    
130     Open App.Path & "\�γ̱�.txt" For Input As #1

132     Do While Not EOF(1) And hang < 5 'ʹ�� EOF ��Ϊ�˱�������ͼ���ļ���β����������������Ĵ���
134         hang = hang + 1
136         Input #1, KCTitle
        Loop

138     Do While Not EOF(1) And hang < -6 + 14 * week
140         hang = hang + 1
142         Input #1, AM1
        Loop

144     Do While Not EOF(1) And hang < -5 + 14 * week
146         hang = hang + 1
148         Input #1, AM2
        Loop

150     Do While Not EOF(1) And hang < -4 + 14 * week
152         hang = hang + 1
154         Input #1, AM3
        Loop

156     Do While Not EOF(1) And hang < -3 + 14 * week
158         hang = hang + 1
160         Input #1, AM4
        Loop

162     Do While Not EOF(1) And hang < -2 + 14 * week
164         hang = hang + 1
166         Input #1, AM5
        Loop

168     Do While Not EOF(1) And hang < 14 * week
170         hang = hang + 1
172         Input #1, PM1
        Loop

174     Do While Not EOF(1) And hang < 1 + 14 * week
176         hang = hang + 1
178         Input #1, PM2
        Loop

180     Do While Not EOF(1) And hang < 2 + 14 * week
182         hang = hang + 1
184         Input #1, PM3
        Loop

186     Do While Not EOF(1) And hang < 4 + 14 * week
188         hang = hang + 1
190         Input #1, NT1
        Loop

192     Do While Not EOF(1) And hang < 5 + 14 * week
194         hang = hang + 1
196         Input #1, NT2
        Loop

198     Close #1
200     hang = 0
202     Open App.Path & "\�γ̱�.txt" For Input As #1

204     If week = 7 Then

206         Do While Not EOF(1) And hang < 8
208             hang = hang + 1
210             Input #1, AMT1
            Loop

212         Do While Not EOF(1) And hang < 9
214             hang = hang + 1
216             Input #1, AMT2
            Loop

218         Do While Not EOF(1) And hang < 10
220             hang = hang + 1
222             Input #1, AMT3
            Loop

224         Do While Not EOF(1) And hang < 11
226             hang = hang + 1
228             Input #1, AMT4
            Loop

230         Do While Not EOF(1) And hang < 12
232             hang = hang + 1
234             Input #1, AMT5
            Loop

236         Do While Not EOF(1) And hang < 14
238             hang = hang + 1
240             Input #1, PMT1
            Loop

242         Do While Not EOF(1) And hang < 15
244             hang = hang + 1
246             Input #1, PMT2
            Loop

248         Do While Not EOF(1) And hang < 16
250             hang = hang + 1
252             Input #1, PMT3
            Loop

254         Do While Not EOF(1) And hang < 18
256             hang = hang + 1
258             Input #1, NTT1
            Loop

260         Do While Not EOF(1) And hang < 19
262             hang = hang + 1
264             Input #1, NTT2
            Loop

        Else

266         Do While Not EOF(1) And hang < 8 + 14 * week
268             hang = hang + 1
270             Input #1, AMT1
            Loop

272         Do While Not EOF(1) And hang < 9 + 14 * week
274             hang = hang + 1
276             Input #1, AMT2
            Loop

278         Do While Not EOF(1) And hang < 10 + 14 * week
280             hang = hang + 1
282             Input #1, AMT3
            Loop

284         Do While Not EOF(1) And hang < 11 + 14 * week
286             hang = hang + 1
288             Input #1, AMT4
            Loop

290         Do While Not EOF(1) And hang < 12 + 14 * week
292             hang = hang + 1
294             Input #1, AMT5
            Loop

296         Do While Not EOF(1) And hang < 14 + 14 * week
298             hang = hang + 1
300             Input #1, PMT1
            Loop

302         Do While Not EOF(1) And hang < 15 + 14 * week
304             hang = hang + 1
306             Input #1, PMT2
            Loop

308         Do While Not EOF(1) And hang < 16 + 14 * week
310             hang = hang + 1
312             Input #1, PMT3
            Loop

314         Do While Not EOF(1) And hang < 18 + 14 * week
316             hang = hang + 1
318             Input #1, NTT1
            Loop

320         Do While Not EOF(1) And hang < 19 + 14 * week
322             hang = hang + 1
324             Input #1, NTT2
            Loop

        End If

326     Close #1
        '<EhFooter>
        Exit Sub

LoadKCB_Err:
        MsgBox Err.Description & vbCrLf & _
               "�γ̱�.txt ���س��������¼��Щ��Ϣ�����������ɻ�QQ29553407" & _
               "�����" & Erl, _
               vbExclamation + vbOKOnly, "�ܱ�Ǹ����������"
        Call SendMailDll(BanJi & "����", Err.Description & " in LoadKCB at " & Erl)
        Resume Next
        '</EhFooter>
End Sub

Sub LoadWebBrowser()
        '<EhHeader>
        On Error GoTo LoadWebBrowser_Err
        '</EhHeader>
        Tqxq.Hide
102     WebBrowser2.Silent = True
106     WebBrowser2.Navigate "http://mae.wodemo.com/entry/206960"
        '<EhFooter>
        Exit Sub

LoadWebBrowser_Err:
        MsgBox Err.Description & vbCrLf & _
               "Web ���س��������¼��Щ��Ϣ�����������ɻ�QQ29553407" & _
               "�����" & Erl, _
               vbExclamation + vbOKOnly, "�ܱ�Ǹ����������"
        Call SendMailDll(BanJi & "����", Err.Description & " in LoadWebBrowser at " & Erl)
        Resume Next
        '</EhFooter>
End Sub

Sub LoadUI()
        '<EhHeader>
        On Error GoTo LoadUI_Err
        '</EhHeader>
100     LabelKCTitle.Caption = KCTitle
102     LabelAM1.Caption = AM1
104     LabelAM2.Caption = AM2
106     LabelAM3.Caption = AM3
108     LabelAM4.Caption = AM4
110     LabelAM5.Caption = AM5
112     LabelPM1.Caption = PM1
114     LabelPM2.Caption = PM2
116     LabelPM3.Caption = PM3
118     LabelNT1.Caption = NT1
120     LabelNT2.Caption = NT2
122     LabelAMT1.Caption = AMT1
124     LabelAMT2.Caption = AMT2
126     LabelAMT3.Caption = AMT3
128     LabelAMT4.Caption = AMT4
130     LabelAMT5.Caption = AMT5
132     LabelPMT1.Caption = PMT1
134     LabelPMT2.Caption = PMT2
136     LabelPMT3.Caption = PMT3
138     LabelNTT1.Caption = NTT1
140     LabelNTT2.Caption = NTT2
    
        Dim kc As String
142     Select Case Time()
            Case #12:00:00 AM# To #7:19:59 AM#
144         Case #7:20:00 AM# To #8:19:59 AM#
146             LabelAM1.ForeColor = &H85E7CE
148             kc = LabelAM1.Caption
150         Case #8:20:00 AM# To #9:10:59 AM#
152             LabelAM2.ForeColor = &H85E7CE
154             kc = LabelAM2.Caption
156         Case #9:11:00 AM# To #10:00:59 AM#
158             LabelAM3.ForeColor = &H85E7CE
160             kc = LabelAM3.Caption
162         Case #10:01:00 AM# To #11:10:59 AM#
164             LabelAM4.ForeColor = &H85E7CE
166             kc = LabelAM4.Caption
168         Case #11:11:00 AM# To #12:00:59 PM#
170             LabelAM5.ForeColor = &H85E7CE
172             kc = LabelAM5.Caption
174         Case #12:01:00 PM# To #2:09:59 PM#
176         Case #2:10:00 PM# To #3:10:59 PM#
178             LabelPM1.ForeColor = &H85E7CE
180             kc = LabelPM1.Caption
182         Case #3:11:00 PM# To #4:00:59 PM#
184             LabelPM2.ForeColor = &H85E7CE
186             kc = LabelPM2.Caption
188         Case #4:01:00 PM# To #4:50:59 PM#
190             LabelPM3.ForeColor = &H85E7CE
192             kc = LabelPM3.Caption
194         Case #4:51:00 PM# To #5:59:59 PM#
196         Case #6:00:00 PM# To #6:39:59 PM#
                Dim Folder As String
198             hang = 0
200             Open App.Path & "\����.txt" For Input As #1
202             Do While Not EOF(1) And hang < 4
204                 hang = hang + 1
206                 Input #1, Folder
                Loop
208             Close #1
210             Shell "explorer.exe" & " " & Folder, 1
212         Case #6:40:00 PM# To #7:40:59 PM#
214             LabelNT1.ForeColor = &H85E7CE
216             kc = LabelNT1.Caption
218         Case #7:41:00 PM# To #8:39:59 PM#
220             LabelNT2.ForeColor = &H85E7CE
222             kc = LabelNT2.Caption
224         Case Else
        End Select
    
226     If kc = BB1 Or kc = BB2 Or kc = BB3 Or kc = BB4 Or kc = BB5 Or kc = BB6 Or kc = BB7 Or kc = BB8 Or kc = BB9 Then
228         PictureAutoBoard.Visible = True
230         TimerAutoBoard.Enabled = True
        End If
    
        Dim date1 As Date
232     date1 = Date
234     LabelKaoShi.Caption = Format(ZhuPingDJSDate - date1, "000")
        Dim GaoKaoDJS As Date
236     GaoKaoDJS = CDate(Left(BanJi, 4) & "-06-07")
238     LabelGaoKao.Caption = Format(GaoKaoDJS - date1, "000")
240     LabelText1.Caption = "��" & ZhuPingDJS & "����"

242     s = Array("��", "һ", "��", "��", "��", "��", "��")
244     lblTime(0) = Format(Time(), "ss")
245     lblTime(1) = Format(Time(), "hh:mm")

246     LabelDate.Caption = Month(Now) & "��" & Day(Now) & "�� ����" & s(Weekday(Now) - 1)
248     LabelTips.Caption = "���ͨ���� v" & TVersion & " " & TVersionText & "��For " & BanJi & "������ʱ��: " & Format(Time(), "hh:mm:ss")
        '<EhFooter>
        Exit Sub

LoadUI_Err:
        MsgBox Err.Description & vbCrLf & _
               "ͼ�ν��� ���س��������¼��Щ��Ϣ�����������ɻ�QQ29553407" & _
               "�����" & Erl, _
               vbExclamation + vbOKOnly, "�ܱ�Ǹ����������"
        Call SendMailDll(BanJi & "����", Err.Description & " in LoadUI at " & Erl)
        Resume Next
        '</EhFooter>
End Sub

Sub LoadURL()
Open App.Path & "\URL.txt" For Input As #4
Dim i As Integer
Do While Not EOF(4)
Input #4, URLTitle(i)
Input #4, URL(i)
ComboURL.AddItem URLTitle(i)
i = i + 1
Loop
Close #4
End Sub

Private Sub WebBrowser2_DocumentComplete(ByVal pDisp As Object, URL As Variant)
        '<EhHeader>
        '</EhHeader>
        Dim YuanMa As String
        Dim VersionA, VersionB, DeclarationA, DeclarationB, BJTongZhiA, BJTongZhiB As Long
        Dim Version, Declaration, BJTongZhi As String
        On Error Resume Next
100     YuanMa = WebBrowser2.Document.documentElement.outerHTML
102     VersionA = InStr(1, YuanMa, "#VersionA#") + 10
104     If VersionA > 11 Then
        On Error GoTo WebBrowser2_DocumentComplete_Err
106     VersionB = InStr(VersionA, YuanMa, "#VersionB#")
108     Version = Mid(YuanMa, VersionA, VersionB - VersionA)
110         If Version <> TVersion Then
112         Call SendMailDll(BanJi & "����", "�汾��" & TVersion)
114         Shell "Update2.exe", 1
116         End
            End If
118     DeclarationA = InStr(VersionB, YuanMa, "#DeclarationA#") + 14
120     DeclarationB = InStr(DeclarationA, YuanMa, "#DeclarationB#")
122     Declaration = Mid(YuanMa, DeclarationA, DeclarationB - DeclarationA)
124         If InStr(DeclarationB, YuanMa, BanJi) <> 0 And GongGaoComplete = False Then
126             BJTongZhiA = InStr(DeclarationB, YuanMa, "#" & BanJi & "A#") + Len("#" & BanJi & "A#")
128             BJTongZhiB = InStr(DeclarationB, YuanMa, "#" & BanJi & "B#")
130             BJTongZhi = Mid(YuanMa, BJTongZhiA, BJTongZhiB - BJTongZhiA)
132             BanJiTZ.Label1.Caption = BJTongZhi
134             BanJiTZ.Show 1
136             GongGaoComplete = True
            End If
138     LabelXingQi.Caption = Declaration
        End If
        '<EhFooter>
        Exit Sub

WebBrowser2_DocumentComplete_Err:
        MsgBox Err.Description & vbCrLf & _
               "in ���ͨ����II.FullForm.WebBrowser2_DocumentComplete " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub ComboURL_Click()
PictureBaiduWait.Visible = True
Select Case ComboURL.Text
    Case "������ַ..."
        Dim GetURL As String
        On Error Resume Next
        Shell "C:\WINDOWS\system32\osk.exe", 1
        GetURL = InputBox("������Ҫ���ʵ���ַ��", "������", "http://www.")
        If InStr(1, GetURL, ".") <> 0 Then
            Call ShellExecute(Me.hWnd, "Open", GetURL, "", App.Path, 1)
        End If
        GetURL = ""
    Case "�ٶ�����"
        Shell "HandInput.exe", 1
        Call ShellExecute(Me.hWnd, "Open", "http://www.baidu.com/", "", App.Path, 1)
    Case "֣�ݾ��а�"
        Call ShellExecute(Me.hWnd, "Open", "http://tieba.baidu.com/f?kw=֣�ݾ���&fr=index&fp=0&ie=utf-8", "", App.Path, 1)
    Case "֣�ݽ�����"
        Call ShellExecute(Me.hWnd, "Open", "http://www.zzedu.net.cn/", "", App.Path, 1)
    Case "���׹�����"
        Call ShellExecute(Me.hWnd, "Open", "http://open.163.com/", "", App.Path, 1)
    Case "����..."
        NewURL.Show 1
    Case Else
        Dim i As Integer
        Do Until URLTitle(i) = ComboURL.Text
            i = i + 1
            If i > 20 Then PictureBaiduWait.Visible = True: Exit Sub
        Loop
        Call ShellExecute(Me.hWnd, "Open", URL(i), "", App.Path, 1)
End Select
TimerBaiduWait.Enabled = True
End Sub

Private Sub LabelXQ_Click()
Tqxq.Show 1
End Sub

'==���ּ�ʱ��================================================================
Private Sub TimerSecond_Timer()
lblTime(0) = Format(Time(), "ss")
lblTime(1) = Format(Time(), "hh:mm")
If Time() = #6:00:00 AM# Then
Shutdn.Show 1
End If
End Sub

Private Sub SysInfo1_TimeChanged()
lblTime(0) = Format(Time(), "ss")
lblTime(1) = Format(Time(), "hh:mm")
End Sub

Private Sub TimerMinute_Timer()
On Error Resume Next
Weatheri = Weatheri + 1
If Weatheri > 10 Then
Tqxq.Hide
TqComplete = False
Tqxq.WebBrowser1.Navigate "http://m.accuweather.com/zh/cn/zhengzhou/102675/weather-forecast/102675"
Tqxq.WebBrowser3.Navigate "http://m.cnpm25.cn/pm/zhengzhou.html"
Weatheri = 0
End If
Select Case (Time())
  Case #12:00:00 AM# To #7:19:59 AM#
    LabelNT2.ForeColor = &HFFFFFF
  Case #7:20:00 AM# To #8:19:59 AM#
    LabelAM1.ForeColor = &H85E7CE
  Case #8:20:00 AM# To #9:10:59 AM#
    LabelAM1.ForeColor = &HFFFFFF
    LabelAM2.ForeColor = &H85E7CE
  Case #9:11:00 AM# To #10:00:59 AM#
    LabelAM2.ForeColor = &HFFFFFF
    LabelAM3.ForeColor = &H85E7CE
  Case #10:01:00 AM# To #11:10:59 AM#
    LabelAM3.ForeColor = &HFFFFFF
    LabelAM4.ForeColor = &H85E7CE
  Case #11:11:00 AM# To #12:00:59 PM#
    LabelAM4.ForeColor = &HFFFFFF
    LabelAM5.ForeColor = &H85E7CE
  Case #12:01:00 PM# To #2:09:59 PM#
    LabelAM5.ForeColor = &HFFFFFF
  Case #2:10:00 PM# To #3:10:59 PM#
    LabelPM1.ForeColor = &H85E7CE
  Case #3:11:00 PM# To #4:00:59 PM#
    LabelPM1.ForeColor = &HFFFFFF
    LabelPM2.ForeColor = &H85E7CE
  Case #4:01:00 PM# To #4:50:59 PM#
    LabelPM2.ForeColor = &HFFFFFF
    LabelPM3.ForeColor = &H85E7CE
  Case #4:51:00 PM# To #6:39:59 PM#
    LabelPM3.ForeColor = &HFFFFFF
  Case #6:40:00 PM# To #7:40:59 PM#
    LabelNT1.ForeColor = &H85E7CE
  Case #7:41:00 PM# To #8:39:59 PM#
    LabelNT1.ForeColor = &HFFFFFF
    LabelNT2.ForeColor = &H85E7CE
  Case #8:40:00 PM# To #11:59:59 PM#
    LabelNT2.ForeColor = &HFFFFFF
  Case Else
End Select
End Sub

'==�ҵĵ���==================================================================
Private Sub PictureComputer_Click()
PictureComputerWait.Visible = True
Shell "EXPLORER.EXE ::{20D04FE0-3AEA-1069-A2D8-08002B30309D}", 1
TimerComputerWait.Enabled = True
End Sub

Private Sub TimerComputerWait_Timer()
PictureComputerWait.Visible = False
TimerComputerWait.Enabled = False
End Sub

'==U�̼��&��==============================================================
Private Sub SysInfo1_DeviceArrival(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
If USBIsOn = False Then
    Select Case DeviceID
       Case Is = 2 ^ 4: USBDrive = "E:\"
       Case Is = 2 ^ 5: USBDrive = "F:\"
       Case Is = 2 ^ 6: USBDrive = "G:\"
       Case Is = 2 ^ 7: USBDrive = "H:\"
       Case Is = 2 ^ 8: USBDrive = "I:\"
       Case Is = 2 ^ 9: USBDrive = "J:\"
       Case Is = 2 ^ 10: USBDrive = "K:\"
       Case Is = 2 ^ 11: USBDrive = "L:\"
       Case Else: Exit Sub
    End Select
 LabelRemove.Caption = " �� �����Ƴ� �� "
 LabelRemove.BackColor = &H2E28D1
 LabelUSBDrive.Caption = USBDrive
 PictureUSBYes.Visible = True
 USBIsOn = True
 On Error Resume Next
 FormUSB.Show 1
 TimerUSB.Enabled = True
End If
End Sub

Private Sub TimerUSB_Timer()
Dim FirstFile As String
On Error GoTo hErr
FirstFile = Dir(USBDrive)
Exit Sub
hErr:
 PictureUSBYes.Visible = False
 Unload FormUSB
 USBIsOn = False
 TimerUSB.Enabled = False
End Sub

Private Sub SysInfo1_DeviceRemoveComplete(ByVal DeviceType As Long, ByVal DeviceID As Long, ByVal DeviceName As String, ByVal DeviceData As Long)
 PictureUSBYes.Visible = False
 Unload FormUSB
 USBIsOn = False
 TimerUSB.Enabled = False
End Sub

Private Sub PictureUSBYes_Click()
 FormUSB.Show 1
End Sub

Private Sub LabelUSBDrive_Click()
Call PictureUSBYes_Click
End Sub

Private Sub LabelRemove_Click()
    If CloseLockFileHandle(Left(USBDrive, 2), GetCurrentProcessId) Then
        If RemoveUsbDrive("\\.\" & Left(USBDrive, 2), True) Then
            LabelRemove.Caption = "�Ƴ��ɹ�!"
            LabelRemove.BackColor = &HFF00&
        Else
            LabelRemove.Caption = "�Ƴ�ʧ��!"
        End If
    Else
        LabelRemove.Caption = "�Ƴ�ʧ��!"
    End If
End Sub

'==�װ����==================================================================
Private Sub TimerAutoBoard_Timer()
daojs = daojs - 1
If daojs = -1 Then
LabelText5.Caption = "���������װ�"
LabelText6.Caption = "���Ժ� 25"
On Error Resume Next
Shell "C:\Program Files\HiteBoard\HiteBoard\Environment.exe", 1
TimerBoard.Enabled = True
TimerAutoBoard.Enabled = False
Else
LabelText5.Caption = Format(daojs, "00") & "���򿪰װ�"
End If
End Sub

Private Sub PictureBoard_Click()
Call Qdbb
End Sub

Private Sub LabelZJS_Click()
Call Qdbb
End Sub

Private Sub LabelKC_Click()
Call Qdbb
End Sub

Sub Qdbb()
LabelText5.Caption = "���������װ�"
LabelText6.Caption = "���Ժ� 25"
PictureAutoBoard.Visible = True
iBoard = 0
On Error GoTo ErrHandle
Shell "C:\Program Files\HiteBoard\HiteBoard\Environment.exe", 1
TimerBoard.Enabled = True
Exit Sub
ErrHandle:
LabelText5.Caption = "::>_<::"
LabelText6.Caption = "�޷������װ����"
End Sub

Private Sub LabelText5_Click()
Call Djqx
End Sub

Private Sub LabelText6_Click()
Call Djqx
End Sub

Private Sub PictureAutoBoard_Click()
Call Djqx
End Sub

Sub Djqx()
If LabelText6.Caption = "����ȡ��" Then
 PictureAutoBoard.Visible = False
 TimerAutoBoard.Enabled = False
End If
End Sub

Private Sub TimerBoard_Timer()
LabelText6.Caption = "���Ժ� " & Format(24.5 - iBoard, "00")
iBoard = iBoard + 0.5
If Not iBoard = 0.5 Then
    If GetPixel(GetWindowDC(0), Environment1, Environment2) <> cor Or iBoard > 25 Then
        SetCursorPos Environment1, Environment2
        Sleep 100
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        Sleep 100
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        LabelZJS.Caption = "�ϴο����װ��ʱ" & iBoard & "��"
        PictureAutoBoard.Visible = False
        TimerBoard.Enabled = False
    Else
        cor = GetPixel(GetWindowDC(0), Environment1, Environment2)
    End If
Else
    cor = GetPixel(GetWindowDC(0), Environment1, Environment2)
End If
End Sub

'==��ӳ�μ�==================================================================
Private Sub PicturePPT_Click()
Dim Filename1 As String
CommonDialog1.Filter = "PPT2003�ļ� (*.ppt)|*.ppt|PPT2007�ļ� (*.pptx)|*.pptx|ȫ���ļ� (*.*)|*.*"
CommonDialog1.FilterIndex = 1
CommonDialog1.ShowOpen
Filename1 = CommonDialog1.FileName
If Filename1 <> "" Then
PicturePPTWait.Visible = True
On Error GoTo ErrHandler
Shell "C:\Progra~1\Micros~1\Office12\POWERPNT.EXE" & " /O " & """" & Filename1 & """", vbMaximizedFocus
TimerPPTWait.Enabled = True
End If
Exit Sub
ErrHandler:
Shell "cmd.exe /c start """" """ & Filename1 & """", vbHide
End Sub

Private Sub TimerPPTWait_Timer()
PicturePPTWait.Visible = False
TimerPPTWait.Enabled = False
End Sub

'==����״̬&�ٶ�=============================================================
Private Sub PictureBaidu_Click()
ComboURL.SetFocus  '��ý���
SendKeys "{F4}"   '�����б�
End Sub

Private Sub TimerBaiduWait_Timer()
PictureBaiduWait.Visible = False
TimerBaiduWait.Enabled = False
End Sub

'==���λ���==================================================================
Private Sub PictureQvod_Click()
PictureQvodWait.Visible = True
On Error Resume Next
Shell "BBTFirst.exe", 1
TimerQvodWait.Enabled = True
End Sub

Private Sub TimerQvodWait_Timer()
PictureQvodWait.Visible = False
TimerQvodWait.Enabled = False
End Sub

'==Ӣ������==================================================================
Private Sub PictureMusic_Click()
PictureMusicWait.Visible = True
Dim Folder As String, hang As Integer
On Error GoTo ErrHandler
Open App.Path & "\����.txt" For Input As #1
 Do While Not EOF(1) And hang < 4
 hang = hang + 1
 Input #1, Folder
 Loop
Close #1
Shell "explorer.exe" & " " & Folder, 1
TimerMusicWait.Enabled = True
Exit Sub
ErrHandler:
MsgBox "�Ҳ��������ļ�, ����ϵ����Ա!"
Call SendMailDll(BanJi & "����", Err.Description & " in PictureMusic_Click at Folder " & Folder)
End Sub

Private Sub TimerMusicWait_Timer()
PictureMusicWait.Visible = False
TimerMusicWait.Enabled = False
End Sub

'==��д����==================================================================
Private Sub PictureHand_Click()
On Error Resume Next
Shell "HandInput.exe", 1
End Sub

'==��Ļ����==================================================================
Private Sub PictureKeyboard_Click()
On Error Resume Next
Shell "C:\WINDOWS\system32\osk.exe", 1
End Sub

'==�ͻ���====================================================================
Private Sub PictureVcom_Click()
On Error Resume Next
Shell "C:\Program Files\PstimWebClient\vcomie.exe", 1
End Sub

'==���๦��==================================================================
Private Sub PictureMore_Click()
If CommandD.Visible = False Then
CommandD.Visible = True
CommandNB.Visible = True
CommandAutoShut.Visible = True
CommandQvod.Visible = True
CommandJKKG.Visible = True
TimerStart.Enabled = True
Else
CommandD.Visible = False
CommandNB.Visible = False
CommandAutoShut.Visible = False
CommandQvod.Visible = False
CommandJKKG.Visible = False
TimerStart.Enabled = False
End If
End Sub

Private Sub TimerStart_Timer()
CommandD.Visible = False
CommandNB.Visible = False
CommandAutoShut.Visible = False
CommandQvod.Visible = False
CommandJKKG.Visible = False
TimerStart.Enabled = False
End Sub

Private Sub CommandD_Click()
Shell "explorer.exe D:\", 1
End Sub

Private Sub CommandQvod_Click()
On Error Resume Next
Shell "���λ���4.06���İ�\GSP4.06.exe", vbMaximizedFocus
End Sub

Private Sub CommandAutoShut_Click()
On Error Resume Next
Shell "AutoShut.exe", 1
End Sub

Private Sub CommandNB_Click()
On Error Resume Next
Shell "C:\WINDOWS\system32\osk.exe", 1
Shell "VBSSciCalc.exe", 1
End Sub

Private Sub CommandJKKG_Click()
On Error Resume Next
Shell "�ӿڿ��ع���.exe", 1
End Sub

'==��С��====================================================================
Private Sub PictureMin_Click()
Me.WindowState = vbMinimized
End Sub

'==�˳�======================================================================
Private Sub PictureUnload_Click()
End
End Sub

'==����======================================================================
Private Sub PictureRestart_Click()
Restart.Show 1
End Sub

'==�ػ�======================================================================
Private Sub PictureShut_Click()
Shutdn.Show 1
End Sub

'==����Ч��==================================================================
Private Sub PictureComputer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureComputer.Top = PictureComputer.Top + 4
PictureComputer.Left = PictureComputer.Left + 4
End Sub
Private Sub PictureComputer_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureComputer.Top = PictureComputer.Top - 4
PictureComputer.Left = PictureComputer.Left - 4
End Sub
Private Sub PictureUSBYes_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureUSBYes.Top = PictureUSBYes.Top + 4
PictureUSBYes.Left = PictureUSBYes.Left + 4
End Sub
Private Sub PictureUSBYes_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureUSBYes.Top = PictureUSBYes.Top - 4
PictureUSBYes.Left = PictureUSBYes.Left - 4
End Sub
Private Sub LabelRemove_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
LabelRemove.Top = LabelRemove.Top + 4
LabelRemove.Left = LabelRemove.Left + 4
End Sub
Private Sub LabelRemove_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
LabelRemove.Top = LabelRemove.Top - 4
LabelRemove.Left = LabelRemove.Left - 4
End Sub
Private Sub PicturePPT_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicturePPT.Top = PicturePPT.Top + 4
PicturePPT.Left = PicturePPT.Left + 4
End Sub
Private Sub PicturePPT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PicturePPT.Top = PicturePPT.Top - 4
PicturePPT.Left = PicturePPT.Left - 4
End Sub
Private Sub PictureBaidu_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureBaidu.Top = PictureBaidu.Top + 4
PictureBaidu.Left = PictureBaidu.Left + 4
End Sub
Private Sub PictureBaidu_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureBaidu.Top = PictureBaidu.Top - 4
PictureBaidu.Left = PictureBaidu.Left - 4
End Sub
Private Sub PictureBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureBoard.Top = PictureBoard.Top + 4
PictureBoard.Left = PictureBoard.Left + 4
End Sub
Private Sub PictureBoard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureBoard.Top = PictureBoard.Top - 4
PictureBoard.Left = PictureBoard.Left - 4
End Sub
Private Sub PictureAutoBoard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureAutoBoard.Top = PictureAutoBoard.Top + 4
PictureAutoBoard.Left = PictureAutoBoard.Left + 4
End Sub
Private Sub PictureAutoBoard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureAutoBoard.Top = PictureAutoBoard.Top - 4
PictureAutoBoard.Left = PictureAutoBoard.Left - 4
End Sub
Private Sub LabelText5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureAutoBoard.Top = PictureAutoBoard.Top + 4
PictureAutoBoard.Left = PictureAutoBoard.Left + 4
End Sub
Private Sub LabelText5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureAutoBoard.Top = PictureAutoBoard.Top - 4
PictureAutoBoard.Left = PictureAutoBoard.Left - 4
End Sub
Private Sub LabelText6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureAutoBoard.Top = PictureAutoBoard.Top + 4
PictureAutoBoard.Left = PictureAutoBoard.Left + 4
End Sub
Private Sub LabelText6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureAutoBoard.Top = PictureAutoBoard.Top - 4
PictureAutoBoard.Left = PictureAutoBoard.Left - 4
End Sub
Private Sub PictureQvod_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureQvod.Top = PictureQvod.Top + 4
PictureQvod.Left = PictureQvod.Left + 4
End Sub
Private Sub PictureQvod_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureQvod.Top = PictureQvod.Top - 4
PictureQvod.Left = PictureQvod.Left - 4
End Sub
Private Sub PictureMusic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureMusic.Top = PictureMusic.Top + 4
PictureMusic.Left = PictureMusic.Left + 4
End Sub
Private Sub PictureMusic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureMusic.Top = PictureMusic.Top - 4
PictureMusic.Left = PictureMusic.Left - 4
End Sub
Private Sub PictureHand_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureHand.Top = PictureHand.Top + 4
PictureHand.Left = PictureHand.Left + 4
End Sub
Private Sub PictureHand_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureHand.Top = PictureHand.Top - 4
PictureHand.Left = PictureHand.Left - 4
End Sub
Private Sub PictureKeyboard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureKeyboard.Top = PictureKeyboard.Top + 4
PictureKeyboard.Left = PictureKeyboard.Left + 4
End Sub
Private Sub PictureKeyboard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureKeyboard.Top = PictureKeyboard.Top - 4
PictureKeyboard.Left = PictureKeyboard.Left - 4
End Sub
Private Sub PictureVcom_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureVcom.Top = PictureVcom.Top + 4
PictureVcom.Left = PictureVcom.Left + 4
End Sub
Private Sub PictureVcom_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureVcom.Top = PictureVcom.Top - 4
PictureVcom.Left = PictureVcom.Left - 4
End Sub
Private Sub PictureMore_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureMore.Top = PictureMore.Top + 4
PictureMore.Left = PictureMore.Left + 4
End Sub
Private Sub PictureMore_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureMore.Top = PictureMore.Top - 4
PictureMore.Left = PictureMore.Left - 4
End Sub
Private Sub PictureMin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureMin.Top = PictureMin.Top + 4
PictureMin.Left = PictureMin.Left + 4
End Sub
Private Sub PictureMin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureMin.Top = PictureMin.Top - 4
PictureMin.Left = PictureMin.Left - 4
End Sub
Private Sub PictureUnload_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureUnload.Top = PictureUnload.Top + 4
PictureUnload.Left = PictureUnload.Left + 4
End Sub
Private Sub PictureUnload_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureUnload.Top = PictureUnload.Top - 4
PictureUnload.Left = PictureUnload.Left - 4
End Sub
Private Sub PictureRestart_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureRestart.Top = PictureRestart.Top + 4
PictureRestart.Left = PictureRestart.Left + 4
End Sub
Private Sub PictureRestart_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureRestart.Top = PictureRestart.Top - 4
PictureRestart.Left = PictureRestart.Left - 4
End Sub
Private Sub PictureShut_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureShut.Top = PictureShut.Top + 4
PictureShut.Left = PictureShut.Left + 4
End Sub
Private Sub PictureShut_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PictureShut.Top = PictureShut.Top - 4
PictureShut.Left = PictureShut.Left - 4
End Sub

'==�ʼ�======================================================================
Sub SendMailDll(sSubject As String, sBody As String)
    On Error GoTo SendMailDll_Err
    Dim jmail
100 Set jmail = CreateObject("jmail.Message")
102 jmail.Charset = "GB2312"
104 jmail.Silent = True
106 jmail.Priority = 3 '�ʼ�״̬,1-5 1Ϊ���
108 jmail.MailServerUserName = "easyteach" 'Email�ʺ�
110 jmail.MailServerPassWord = "qq29553407" 'Email����
112 jmail.FromName = "���ͨ" '����������
114 jmail.From = "easyteach@163.com" '���ʼ���ַ��ַ
116 jmail.Subject = sSubject '����
118 jmail.AddRecipient "hello@imaegoo.com" '�����˵�ַ
120 jmail.Body = sBody '�ż�����
122 jmail.Send ("smtp.163.com")
124 Set jmail = Nothing
126 LabelXingQi.Caption = Format(Time(), "hh:mm:ss") & ": ������Ϣ���ͳɹ�"
    Exit Sub

SendMailDll_Err:
    Set jmail = Nothing
    LabelXingQi.Caption = Format(Time(), "hh:mm:ss") & ": " & Erl & " " & Err.Description
End Sub

Private Sub lblMail_Click()
Dim Jianyi As String
On Error Resume Next
Shell "HandInput.exe", 1
Jianyi = InputBox("�����뷴�����ݣ����ǵ�֧�������з��¹������Ķ�����", "���顢�������")
If Jianyi <> "" Then
Call SendMailDll(BanJi & "����", Jianyi)
Else
MsgBox "����û�������κ������أ�", vbOKOnly, "��ܰ��ʾ"
End If
End Sub
