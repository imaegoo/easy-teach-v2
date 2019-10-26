VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Tqxq 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "天气详情"
   ClientHeight    =   7200
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9600
   Icon            =   "Tqxq.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      Caption         =   "空气质量"
      Height          =   7095
      Left            =   4860
      TabIndex        =   2
      Top             =   60
      Width           =   4695
      Begin SHDocVwCtl.WebBrowser WebBrowser3 
         Height          =   6735
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   4455
         ExtentX         =   7858
         ExtentY         =   11880
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
         Location        =   ""
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "天气预报"
      Height          =   7095
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   4695
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   6735
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4455
         ExtentX         =   7858
         ExtentY         =   11880
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
         Location        =   ""
      End
   End
End
Attribute VB_Name = "Tqxq"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
WebBrowser1.Silent = True
WebBrowser3.Silent = True
WebBrowser1.Navigate "http://m.accuweather.com/zh/cn/zhengzhou/102675/weather-forecast/102675"
WebBrowser3.Navigate "http://m.cnpm25.cn/pm/zhengzhou.html"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = 1
If Me.Visible = True Then
Me.Visible = False
Else
Unload Me
End If
End Sub

Private Sub WebBrowser1_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Dim Tqym     As String
    Dim TqLen1   As Integer, TqLen2 As Double, TqLen3 As Integer, TqLen4 As Integer, TqLen5 As Integer, TqLen6 As Integer, TqLen7 As Integer, TqLen8 As Integer, TqLen9 As Integer, TqLen10 As Integer, TqLenQwMax As Integer, TqLenQwMax2 As Integer
    Dim Tq       As String, Qw As String, QwMax As String, Js As String, Tips As String

    Dim TqMrLen1 As Integer, TqMrLen2 As Double, TqMrLen3 As Integer, TqMrLen4 As Integer, TqMrLen5 As Integer, TqMrLen6 As Integer, TqMrLen7 As Integer, TqMrLen8 As Integer, TqMrLen9 As Integer, TqMrLen10 As Integer, TqMrLenQwMax As Integer, TqLenMrQwMax2 As Integer
    Dim MrTq     As String, MrQw As String, MrQwMax As String

    On Error Resume Next
    If TqComplete = True Then Exit Sub

    Tqym = WebBrowser1.Document.documentElement.outerHTML
    TqLen1 = InStr(1, Tqym, "<H2>未来", 0)

    If TqLen1 > 10 Then
        FullForm.PictureBaidu.Visible = True
    Else
        FullForm.PictureBaidu.Visible = False
    End If

    TqLenQwMax = InStr(TqLen1, Tqym, "<DT>高温", 0) + 13

    If TqLenQwMax - TqLen1 < 700 Then
        TqLenQwMax2 = InStr(TqLen1, Tqym, "<SPAN ", 0)
        QwMax = Mid(Tqym, TqLenQwMax, TqLenQwMax2 - TqLenQwMax)
    End If

    TqLen2 = InStr(TqLen1, Tqym, "<SPAN class=lo>", 0) + 15
    TqLen3 = InStr(TqLen2, Tqym, "<SPAN ", 0)
    TqLen4 = InStr(TqLen3, Tqym, "<DD>", 0) + 4
    TqLen5 = InStr(TqLen4, Tqym, " ", 0)
    TqLen6 = InStr(TqLen5, Tqym, "</SPAN>", 0) + 8
    TqLen7 = InStr(TqLen6, Tqym, "%", 0) + 1
    TqLen8 = InStr(TqLen7, Tqym, "<DD class=wx-alarm><", 0) + 30
    TqLen9 = InStr(TqLen8, Tqym, ">", 0) + 1
    TqLen10 = InStr(TqLen9, Tqym, " </A>", 0)
    Qw = Mid(Tqym, TqLen2, TqLen3 - TqLen2)

    If QwMax <> "" Then
        Qw = Qw & "~" & QwMax
    End If

    TqLenMrQwMax = InStr(TqLen7, Tqym, "<DT>高温", 0) + 13
    TqLenMrQwMax2 = InStr(TqLen7, Tqym, "<SPAN ", 0)
    MrQwMax = Mid(Tqym, TqLenMrQwMax, TqLenMrQwMax2 - TqLenMrQwMax)
    TqMrLen2 = InStr(TqLen7, Tqym, "<SPAN class=lo>", 0) + 15
    TqMrLen3 = InStr(TqMrLen2, Tqym, "<SPAN ", 0)
    TqMrLen4 = InStr(TqMrLen3, Tqym, "<DD>", 0) + 4
    TqMrLen5 = InStr(TqMrLen4, Tqym, " ", 0)
    TqMrLen6 = InStr(TqMrLen5, Tqym, "</SPAN>", 0) + 8
    TqMrLen7 = InStr(TqMrLen6, Tqym, "%", 0) + 1
    TqMrLen8 = InStr(TqMrLen7, Tqym, "<DD class=wx-alarm><", 0) + 30
    TqMrLen9 = InStr(TqMrLen8, Tqym, ">", 0) + 1
    TqMrLen10 = InStr(TqMrLen9, Tqym, " </A>", 0)
    MrQw = Mid(Tqym, TqMrLen2, TqMrLen3 - TqMrLen2)
    MrQw = MrQw & "~" & MrQwMax
    MrTq = Mid(Tqym, TqMrLen4, TqMrLen5 - TqMrLen4)

    Tq = Mid(Tqym, TqLen4, TqLen5 - TqLen4)
    Js = Mid(Tqym, TqLen6, TqLen7 - TqLen6)
    Tips = Mid(Tqym, TqLen9, TqLen10 - TqLen9)

    If Tips = "" Then
        Tips = "无天气预警。"
    End If

    If Tq <> "" And Qw <> "" Then
        FullForm.lblStart(1).Caption = Tq & " " & Qw & "℃"
        FullForm.lblStart(2).Caption = "明日：" & MrTq & " " & MrQw & "℃"
        FullForm.LabelAbout.Caption = "今日降水" & Js & "。" & Tips
        TqComplete = True

        '=================================背景自动换==================================
        If Dir(App.Path & "\天气背景\") <> "" Then
            If InStr(1, Tq, "热") <> 0 Then
                If TqBj <> 1 Then
                    TqBj = 1
                    FullForm.Picture = LoadPicture(App.Path & "\天气背景\热.jpg")
                End If

                Exit Sub
            End If

            If InStr(1, Tq, "冷") <> 0 Then
                If TqBj <> 2 Then
                    TqBj = 2
                    FullForm.Picture = LoadPicture(App.Path & "\天气背景\冷.jpg")
                End If

                Exit Sub
            End If

            If InStr(1, Tq, "阳光") <> 0 Then
                If TqBj <> 3 Then
                    TqBj = 3
                    FullForm.Picture = LoadPicture(App.Path & "\天气背景\阳光.jpg")
                End If

                Exit Sub
            End If

            If InStr(1, Tq, "晴") <> 0 Then
                If TqBj <> 4 Then
                    TqBj = 4
                    FullForm.Picture = LoadPicture(App.Path & "\天气背景\晴.jpg")
                End If

                Exit Sub
            End If

            If InStr(1, Tq, "云量") <> 0 Then
                If TqBj <> 5 Then
                    TqBj = 5
                    FullForm.Picture = LoadPicture(App.Path & "\天气背景\云量.jpg")
                End If

                Exit Sub
            End If

            If InStr(1, Tq, "云") <> 0 Then
                If TqBj <> 6 Then
                    TqBj = 6
                    FullForm.Picture = LoadPicture(App.Path & "\天气背景\云.jpg")
                End If

                Exit Sub
            End If

            If InStr(1, Tq, "阴") <> 0 Then
                If TqBj <> 7 Then
                    TqBj = 7
                    FullForm.Picture = LoadPicture(App.Path & "\天气背景\阴.jpg")
                End If

                Exit Sub
            End If

            If InStr(1, Tq, "雾") <> 0 Then
                If TqBj <> 8 Then
                    TqBj = 8
                    FullForm.Picture = LoadPicture(App.Path & "\天气背景\雾.jpg")
                End If

                Exit Sub
            End If

            If InStr(1, Tq, "雷") <> 0 Then
                If TqBj <> 9 Then
                    TqBj = 9
                    FullForm.Picture = LoadPicture(App.Path & "\天气背景\雷.jpg")
                End If

                Exit Sub
            End If

            If InStr(1, Tq, "雨") <> 0 Then
                If TqBj <> 10 Then
                    TqBj = 10
                    FullForm.Picture = LoadPicture(App.Path & "\天气背景\雨.jpg")
                End If

                Exit Sub
            End If

            If InStr(1, Tq, "雪") <> 0 Then
                If TqBj <> 11 Then
                    TqBj = 11
                    FullForm.Picture = LoadPicture(App.Path & "\天气背景\雪.jpg")
                    LabelTips.ForeColor = &H0&
                    lblMail.ForeColor = &H0&
                    LabelXQ.ForeColor = &H0&
                End If

                Exit Sub
            End If

            If InStr(1, Tq, "风") <> 0 Then
                If TqBj <> 12 Then
                    TqBj = 12
                    FullForm.Picture = LoadPicture(App.Path & "\天气背景\风.jpg")
                End If

                Exit Sub
            End If
        End If
    End If

End Sub

Private Sub WebBrowser3_DocumentComplete(ByVal pDisp As Object, URL As Variant)
        '<EhHeader>
        On Error GoTo WebBrowser3_DocumentComplete_Err
        '</EhHeader>
    Dim Pmym As String
    Dim PmLen1 As Integer, PmLen2 As Double, PmLen3 As Integer, PmLen4 As Integer
    Dim Pm As String, PmStr As String
100 Pmym = WebBrowser3.Document.documentElement.outerHTML
102 If InStr(1, Pmym, "郑州") <> 0 Then
104     PmLen1 = InStr(InStr(1, Pmym, "郑州市实时空气质量指数"), Pmym, "FONT-WEIGHT:", vbTextCompare) + 19
106     PmLen2 = InStr(PmLen1, Pmym, "<SPAN", vbTextCompare)
108     PmLen3 = InStr(PmLen2, Pmym, "../img/main/", vbTextCompare) + Len("../img/main/")
110     PmLen4 = InStr(PmLen3, Pmym, """", vbTextCompare)
112     Pm = Mid(Pmym, PmLen1, PmLen2 - PmLen1)
114     If InStr(1, Pm, ">") <> 0 Then
116         Pm = Right(Pm, Len(Pm) - InStr(1, Pm, ">"))
        End If
118     PmStr = Mid(Pmym, PmLen3, PmLen4 - PmLen3)
120     Select Case PmStr
           Case "you.png": PmStr = "优"
122        Case "liang.png": PmStr = "良"
124        Case "qing.png": PmStr = "轻度污染"
126        Case "zhong.png": PmStr = "中度污染"
128        Case "zhongdu.png": PmStr = "重度污染"
130        Case "yan.png": PmStr = "严重污染"
132        Case Else: PmStr = "未知级别"
        End Select
134     FullForm.LabelPMZS.Caption = "空气质量:　" & Pm
136     FullForm.LabelPMZS2.Caption = PmStr

138     If Pm < 350 Then
140         FullForm.PicturePM.Width = Int(1024 * (Pm / 350))
142         FullForm.LabelPMZS2.Left = 15240 * (Pm / 350) - 960
        Else
144         FullForm.PicturePM.Width = 1023
146         FullForm.LabelPMZS2.Caption = "…爆表了"
148         FullForm.LabelPMZS2.Left = 14280
        End If
150     FullForm.PicturePM.Visible = True
    End If
        '<EhFooter>
        Exit Sub

WebBrowser3_DocumentComplete_Err:
        FullForm.LabelXingQi.Caption = Err.Description & _
               "in 空气质量模块 " & _
               "at " & Erl
        Resume Next
        '</EhFooter>
End Sub
