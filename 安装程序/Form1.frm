VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form BBTUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "֣�ݾ��а��ͨ���� ���߰�װ"
   ClientHeight    =   4935
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   6855
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   6855
   StartUpPosition =   1  '����������
   Begin VB.CheckBox Check1 
      Caption         =   "�����Ķ���������"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   4500
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H002E2DC3&
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   9
      Text            =   "Form1.frx":4322
      Top             =   120
      Width           =   2655
   End
   Begin VB.CommandButton Command1 
      Caption         =   "��ʼ����"
      Enabled         =   0   'False
      Height          =   360
      Left            =   4560
      TabIndex        =   8
      Top             =   4440
      Width           =   990
   End
   Begin VB.Frame Frame2 
      Caption         =   "������Ϣ"
      Height          =   2415
      Left            =   2880
      TabIndex        =   6
      Top             =   120
      Width           =   3855
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���Ķ�˵���ٵ������ʼ���ء�"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H002E2DC3&
         Height          =   240
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   3570
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "������Ϣ"
      Height          =   1215
      Left            =   2880
      TabIndex        =   2
      Top             =   2640
      Width           =   3855
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "׼��"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   120
         TabIndex        =   4
         Top             =   420
         Width           =   360
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����״̬��"
         ForeColor       =   &H8000000D&
         Height          =   180
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   900
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "ֹͣ����"
      Enabled         =   0   'False
      Height          =   360
      Left            =   5640
      TabIndex        =   0
      Top             =   4440
      Width           =   1020
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   255
      Left            =   6360
      TabIndex        =   5
      Top             =   240
      Width           =   255
      ExtentX         =   450
      ExtentY         =   450
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
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   2880
      TabIndex        =   1
      Top             =   3960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "BBTUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private WithEvents Cdf As Cls_DownLoad
Attribute Cdf.VB_VarHelpID = -1

Dim UpURL As String

Private Sub Check1_Click()
If Check1.Value = 1 Then
Command1.Enabled = True
Else
Command1.Enabled = False
End If
End Sub

Private Sub Command1_Click()
    WebBrowser1.Silent = True
    WebBrowser1.Navigate "http://mae.wodemo.com/entry/206960"
    Label1.Caption = "���ڶ�ȡ...." & vbCrLf & "�����ʱ��δ��Ӧ��������"
    Check1.Enabled = False
    Command2.Enabled = True
    Command1.Enabled = False
End Sub

Private Sub Form_Load()
    If Dir("C:\PROGRA~1\VCOMTO~1\") = "" Then
        MsgBox "���ڰ��ͨ�ϰ�װ�������", 16, "��ʾ"
        End
    End If
    Set Cdf = New Cls_DownLoad
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
    If MsgBox("���Ҫ�˳�����ô��", 1 + 48 + 256, "����") = vbOK Then
    Cdf.DLFileStop 'ֹͣ����
    End
    End If
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Dim YuanMa As String
    Dim VersionA, VersionB, URLA, URLB, UpdateA, UpdateB As Long
    Dim Version, Update As String
    On Error Resume Next
    YuanMa = WebBrowser1.Document.documentElement.outerHTML
    VersionA = InStr(1, YuanMa, "#VersionA#") + 10

    If VersionA > 11 Then
        VersionB = InStr(VersionA, YuanMa, "#VersionB#")
        Version = Mid(YuanMa, VersionA, VersionB - VersionA)
        URLA = InStr(VersionB, YuanMa, "#FirstInstA#") + 12
        URLB = InStr(URLA, YuanMa, "#FirstInstB#")
        UpURL = Mid(YuanMa, URLA, URLB - URLA)
        UpdateA = InStr(URLB, YuanMa, "#UpdateA#") + 9
        UpdateB = InStr(UpdateA, YuanMa, "#UpdateB#")
        Update = vbCrLf & Replace(Mid(YuanMa, UpdateA, UpdateB - UpdateA), "<BR>", vbCrLf)
        Dim sTmp$
        sTmp = "���°汾��" & Version
        sTmp = sTmp & vbCrLf & "���ص�ַ��" & UpURL
        sTmp = sTmp & vbCrLf & "������־��" & Update
        Label1 = sTmp
        If Cdf.DLFile(UpURL, App.Path & "\FirstInst.exe", 10 * 1000) Then
            Shell App.Path & "\FirstInst.exe", 1
            End
        End If
    End If

End Sub

Private Sub Cdf_entDLFileDowning(sRemoteURL As String, lDownLoaded As Long, lFilesize As Long, lSpeed As Long)
    Dim sTmp$, lTmp&
    If lSpeed = 0 Then lSpeed = 10
    If lFilesize > 0 Then lTmp = (lDownLoaded / lFilesize * 100)
    sTmp = "���ؽ��ȣ� " & Format(lTmp, "0.0") & " %"
    ProgressBar1.Value = lTmp
    sTmp = sTmp & vbCrLf & "�ļ���С�� " & Format(lFilesize / 1024, "0.00") & " KB"
    sTmp = sTmp & vbCrLf & "�Ѿ����أ� " & Format(lDownLoaded / 1024, "0.00") & " KB"
    sTmp = sTmp & vbCrLf & "�����ٶȣ� " & Format(lSpeed / 1024, "0.00") & " KB/s"
    Label4 = sTmp
End Sub

Private Sub Cdf_entDLFileStatus(TmpState As eDL_Status)
    Select Case TmpState
        Case 1
            sConnectStauts = "���ӷ�����..."
        Case 2
            sConnectStauts = "��������..."
        Case 3
            sConnectStauts = "��ȡԶ���ļ���Ϣ..."
        Case 4
            sConnectStauts = "��������..."
        Case 5
            sConnectStauts = "ֹͣ����"
        Case 6
            sConnectStauts = "�������"
        Case 7
            sConnectStauts = "���ӷ�����ʧ��"
        Case 8
            sConnectStauts = "��������ʧ��"
        Case 9
            sConnectStauts = "���ӷ�����"
        Case Else
            sConnectStauts = "���ر���ֹ"
    End Select
    Label5 = "����״̬�� " & sConnectStauts
End Sub

Private Sub Command2_Click()
    If MsgBox("���Ҫֹͣô��", 1 + 48 + 256, "����") = vbOK Then
    Cdf.DLFileStop 'ֹͣ����
    Check1.Enabled = True
    Command1.Enabled = True
    Command2.Enabled = False
    End If
End Sub
