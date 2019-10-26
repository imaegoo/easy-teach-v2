VERSION 5.00
Begin VB.Form FormUSB 
   BackColor       =   &H00D9A108&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "发现新硬件"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8040
   FillColor       =   &H80000005&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FormUSB.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   402
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   536
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      Height          =   795
      Left            =   6840
      Picture         =   "FormUSB.frx":4322
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   360
      Width           =   795
   End
   Begin VB.Label LabelEEE 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   ">>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   10.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   300
      TabIndex        =   8
      Top             =   840
      Width           =   6300
   End
   Begin VB.Label lblRemove 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "安全删除硬件"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   300
      TabIndex        =   6
      Top             =   4800
      Width           =   3150
   End
   Begin VB.Label lblUSB2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "当前盘符 N/A"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   5100
      TabIndex        =   5
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label lblPpt 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "打开PPT"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   300
      TabIndex        =   4
      Top             =   2100
      Width           =   1980
   End
   Begin VB.Label lblNothing 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   300
      TabIndex        =   3
      Top             =   3900
      Width           =   1050
   End
   Begin VB.Label lbl2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "直接打开U盘"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   690
      Left            =   300
      TabIndex        =   2
      Top             =   3000
      Width           =   3015
   End
   Begin VB.Label lblU2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "请选择 ..."
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   420
      Left            =   300
      TabIndex        =   1
      Top             =   1410
      Width           =   1260
   End
   Begin VB.Label lblU 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "已连接U盘"
      BeginProperty Font 
         Name            =   "微软雅黑"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   465
      Left            =   300
      TabIndex        =   0
      Top             =   360
      Width           =   1710
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H001B6AF9&
      FillStyle       =   0  'Solid
      Height          =   1335
      Left            =   0
      Top             =   0
      Width           =   8055
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H00B39638&
      FillStyle       =   0  'Solid
      Height          =   615
      Left            =   0
      Top             =   1320
      Width           =   8055
   End
End
Attribute VB_Name = "FormUSB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PPTFileName As String, PPTFileMax As String, PPTFileDate As Date

Private Sub Form_Load()
    On Error Resume Next
    Dim i As Integer
    lblUSB2.Caption = "当前盘符 " & USBDrive
    PPTFileName = USBDrive & Dir(USBDrive & "*.ppt")
    If Len(PPTFileName) > 3 Then
        Do
            If FileDateTime(PPTFileName) > PPTFileDate Then
                PPTFileDate = FileDateTime(PPTFileName)
                PPTFileMax = PPTFileName
            End If
            PPTFileName = USBDrive & Dir
            i = i + 1
        Loop Until Len(PPTFileName) = 3 Or i > 100
        i = 0
        lblPpt.Caption = "打开PPT: " & Mid(PPTFileMax, 4, InStr(4, PPTFileMax, ".ppt") - 4)
    Else
        lblPpt.Caption = "U盘中无课件"
    End If
End Sub

Private Sub lblPpt_Click()
If lblPpt.Caption = "U盘中无课件" Then
    lblPpt.ForeColor = &H2E28D1
Else
    Shell "cmd.exe /c start """" """ & PPTFileMax & """", vbHide
    Unload Me
End If
End Sub

Private Sub lbl2_Click()
Shell "explorer.exe " & USBDrive, 1
Unload Me
End Sub

Private Sub lblNothing_Click()
Unload Me
End Sub

Private Sub lblRemove_Click()
        '<EhHeader>
        On Error GoTo lblRemove_Click_Err
        '</EhHeader>
        lblRemove.Caption = "请稍候..."
100     If CloseLockFileHandle(Left(USBDrive, 2), GetCurrentProcessId) Then
102         If RemoveUsbDrive("\\.\" & Left(USBDrive, 2), True) Then
104             FullForm.LabelRemove.Caption = "移除成功!"
106             FullForm.LabelRemove.BackColor = &HFF00&
108             Unload Me
            Else
110             lblRemove.Caption = "点击重试"
            End If

        Else
112         lblRemove.Caption = "点击重试"
        End If

        '<EhFooter>
        Exit Sub

lblRemove_Click_Err:
        MsgBox Err.Description & vbCrLf & _
               "in 班班通助手II.FormUSB.lblRemove_Click " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub
