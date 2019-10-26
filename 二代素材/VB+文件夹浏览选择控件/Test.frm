VERSION 5.00
Begin VB.Form fTest 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "文件夹显示控件示例"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8565
   Icon            =   "Test.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5700
   ScaleWidth      =   8565
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "属性 (默认)"
      Height          =   3375
      Left            =   4920
      TabIndex        =   4
      Top             =   1800
      Width           =   3015
      Begin VB.CheckBox chkHasLines 
         BackColor       =   &H00FFFFFF&
         Caption         =   "虚线"
         Height          =   285
         Left            =   720
         TabIndex        =   9
         Top             =   1050
         Value           =   1  'Checked
         Width           =   1440
      End
      Begin VB.CheckBox chkHasButtons 
         BackColor       =   &H00FFFFFF&
         Caption         =   "加号"
         Height          =   285
         Left            =   705
         TabIndex        =   8
         Top             =   645
         Value           =   1  'Checked
         Width           =   1440
      End
      Begin VB.CheckBox chkHideSelection 
         BackColor       =   &H00FFFFFF&
         Caption         =   "隐藏选择"
         Height          =   285
         Left            =   705
         TabIndex        =   7
         Top             =   1455
         Width           =   1440
      End
      Begin VB.CheckBox chkSingleExpand 
         BackColor       =   &H00FFFFFF&
         Caption         =   "只展开一个"
         Height          =   285
         Left            =   705
         TabIndex        =   6
         Top             =   1860
         Width           =   2085
      End
      Begin VB.CheckBox chkTrackSelect 
         BackColor       =   &H00FFFFFF&
         Caption         =   "热跟踪"
         Height          =   285
         Left            =   705
         TabIndex        =   5
         Top             =   2265
         Width           =   1440
      End
   End
   Begin Proyecto1.ucFolderView ucFolderView 
      Height          =   5325
      Left            =   150
      TabIndex        =   3
      Top             =   180
      Width           =   4080
      _ExtentX        =   7197
      _ExtentY        =   9393
   End
   Begin VB.TextBox txtPath 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   4560
      TabIndex        =   1
      Top             =   480
      Width           =   3885
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "应用"
      Default         =   -1  'True
      Height          =   435
      Left            =   7245
      TabIndex        =   2
      Top             =   825
      Width           =   1095
   End
   Begin VB.Label lblPath 
      BackColor       =   &H00FFFFFF&
      Caption         =   "改变的路径："
      Height          =   270
      Left            =   4455
      TabIndex        =   0
      Top             =   165
      Width           =   1815
   End
End
Attribute VB_Name = "fTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'                    ╭════════════════╮
'                    ║      E动天下—VB专业源码网     ║
'  ╭════════┤      网站:http://www.2e3.org   ├═════════╮
'  ║                ╰════════════════╯                  ║
'  ║                            人人为我，我为人人                        ║
'  ║                                                                      ║
'  ║        E动天下vb专业源码网汉化收藏整理                               ║
'  ║                                                                      ║
'  ║        网    站：http://www.2e3.org/                                 ║
 ' ║                                                                      ║
  '║        e-mail  ：ffwl2002@126.com                                    ║
  '║                                                                      ║
'  ║        QQ      ：83892778                                            ║
 ' ║                                                                      ║
  '║    如果您有新的、好的代码可以提供给E动天下上发布，让大家学习哦!      ║
  '║                                                                      ║
  '║                                                                      ║
  '║----------------------------------------------------------------------║
  '║                                                                      ║
 ' ║              ╭════════════════════╮            ║
 ' ║              ║                                        ║            ║
  '╰═══════┤  E动天下—VB专业源码网（www.2e3.org）  ├══════╯
   '               ║                                        ║
    '              ╰════════════════════╯
Option Explicit



Private Sub Form_Load()
Shell "C:\Program Files\Internet Explorer\iexplore.exe http://www.2e3.org"
    Call ucFolderView.Initialize
End Sub


Private Sub ucFolderView_ChangeBefore(ByVal NewPath As String, Cancel As Boolean)
    
    '-- 这里可以检查路径...
End Sub

Private Sub ucFolderView_ChangeAfter(ByVal OldPath As String)

    txtPath.Text = ucFolderView.Path
    txtPath.SelStart = Len(txtPath.Text)
End Sub

Private Sub cmdApply_Click()
    
    ucFolderView.Path = txtPath.Text
End Sub

'//

Private Sub chkHasButtons_Click()
    ucFolderView.HasButtons = CBool(chkHasButtons)
End Sub

Private Sub chkHasLines_Click()
    ucFolderView.HasLines = CBool(chkHasLines)
End Sub

Private Sub chkHideSelection_Click()
    ucFolderView.HideSelection = CBool(chkHideSelection)
End Sub

Private Sub chkSingleExpand_Click()
    ucFolderView.SingleExpand = CBool(chkSingleExpand)
End Sub

Private Sub chkTrackSelect_Click()
    ucFolderView.TrackSelect = CBool(chkTrackSelect)
End Sub


