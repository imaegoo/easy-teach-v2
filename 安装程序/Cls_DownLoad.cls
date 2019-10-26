VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Cls_DownLoad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Download by http://www.codefans.net
'****************************************************************************
'开源世界，有你更加精彩
'作者博客：http://www.80ht.cn
'e-mail  ：admin@80ht.cn
'尊重作者劳动成功，转载请保留
'****************************************************************************
'(DLFile)
Private Declare Function InternetOpen Lib "wininet.dll" Alias "InternetOpenA" (ByVal sAgent As String, ByVal lAccessType As Long, ByVal sProxyName As String, ByVal sProxyBypass As String, ByVal lFlags As Long) As Long
Private Declare Function InternetOpenUrl Lib "wininet.dll" Alias "InternetOpenUrlA" (ByVal hOpen As Long, ByVal surl As String, ByVal sHeaders As String, ByVal lLength As Long, ByVal lFlags As Long, ByVal lContext As Long) As Long

Private Declare Function InternetReadFile Lib "wininet.dll" (ByVal hFile As Long, ByVal sBuffer As String, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer
Private Declare Function InternetReadFileByte Lib "wininet.dll" Alias "InternetReadFile" (ByVal hFile As Long, ByRef sBuffer As Byte, ByVal lNumBytesToRead As Long, lNumberOfBytesRead As Long) As Integer

Private Declare Function InternetCloseHandle Lib "wininet.dll" (ByVal hInet As Long) As Integer
Private Declare Function HttpQueryInfo Lib "wininet.dll" Alias "HttpQueryInfoA" (ByVal hHttpRequest As Long, ByVal lInfoLevel As Long, ByRef sBuffer As Any, ByRef lBufferLength As Long, ByRef lIndex As Long) As Integer
Const INTERNET_OPEN_TYPE_PRECONFIG = 0
Const INTERNET_FLAG_RELOAD = &H80000000
Const INTERNET_FLAG_NO_CACHE_WRITE = &H4000000
Const HTTP_QUERY_CONTENT_LENGTH = 5
Const scUserAgent = "VB Http"
Const BufferSizeA& = 512
Const BufferSizeP& = 512
Public Event entDLFileDowning(sRemoteURL$, lDownLoaded&, lFilesize&, lSpeed&)
Public Event entDLFileStatus(TmpState As eDL_Status)
Public Enum eDL_Status
    s1 = 1  ' 连接服务器
    s2 = 2  ' 发送请求
    s3 = 3  ' 获取远程文件信息
    s4 = 4  ' 开始接收数据
    s5 = 5  ' 停止下载
    s6 = 6  ' 下载完成
    e1 = 7  ' 连接服务器失败
    e2 = 8  ' 发送请求失败
    e3 = 9  ' 下载被中止
End Enum
Private bolStop As Boolean
'#####################################################################################################################################

 '(DLFile)
'**************************************************************************************
Public Function DLFile&(sRemoteURL$, sLocalFilePath$, Optional lTimeOut& = 5000)
    Dim hInternetOpen&, hInternetOpenUrl&, hHttpQueryInfo&, StartTime&, TimeConsuming&, lSpeed&, _
    sBuffer As String * BufferSizeA, lNumberOfBytesRead&, Flng%, TotalReadSize&, bRet&, ReadBuffer() As Byte
    RaiseEvent entDLFileStatus(s1)
    hInternetOpen = InternetOpen(scUserAgent, INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, vbNullString, 0)
    If hInternetOpen Then
        RaiseEvent entDLFileStatus(s2)
        hInternetOpenUrl = InternetOpenUrl(hInternetOpen, sRemoteURL, vbNullString, 0, INTERNET_FLAG_RELOAD, 0)
        If hInternetOpenUrl Then
            RaiseEvent entDLFileStatus(s3)
            hHttpQueryInfo = HttpQueryInfo(hInternetOpenUrl, HTTP_QUERY_CONTENT_LENGTH, ByVal sBuffer, Len(sBuffer), 0)
            sBuffer = IIf(hHttpQueryInfo, Left(sBuffer, Len(sBuffer)), "0")
            ReDim ReadBuffer(BufferSizeA)
            StartTime = Timer
            Flng = FreeFile
            RaiseEvent entDLFileStatus(s4)
            Open sLocalFilePath For Binary Access Write Lock Read As #Flng
                Do
                    bRet = InternetReadFileByte(hInternetOpenUrl, ReadBuffer(0), BufferSizeA, lNumberOfBytesRead)
                    If bRet And lNumberOfBytesRead > 0 Then
                        ReDim Preserve ReadBuffer(0 To lNumberOfBytesRead - 1)
                        Put #Flng, , ReadBuffer
                        TotalReadSize = TotalReadSize + lNumberOfBytesRead
                        TimeConsuming = Timer - StartTime
                        If TimeConsuming > 0 Then
                            lSpeed = TotalReadSize / TimeConsuming
                        Else
                            lSpeed = 1025
                        End If
                        RaiseEvent entDLFileDowning(sRemoteURL, TotalReadSize, CLng(sBuffer), lSpeed)
                    Else
                        RaiseEvent entDLFileStatus(e3)
                        Exit Do
                    End If
                    If bolStop Then
                        bolStop = False
                        RaiseEvent entDLFileStatus(s5)
                        Exit Do
                    End If
                    DoEvents
                Loop
            Close #Flng
            Erase ReadBuffer
            InternetCloseHandle hInternetOpenUrl
        Else
            RaiseEvent entDLFileStatus(e2)
        End If
        InternetCloseHandle hInternetOpen
        RaiseEvent entDLFileStatus(s6)
        DLFile = TotalReadSize
    Else
        RaiseEvent entDLFileStatus(e1)
    End If
End Function

Public Function DLFileStop() As Boolean
    bolStop = True
End Function
