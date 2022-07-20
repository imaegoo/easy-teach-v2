Attribute VB_Name = "ModuleINI"
Option Explicit
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

'读取ini
'Option1.Value = MyGetSetting("项", "名称", "默认值")
Public Function MyGetSetting(Section As String, KeyName As String, DefaultValue As String) As String
Dim X As Long
Dim Holder As String * 255
If Right(App.Path, 2) = ":\" Then
    X = GetPrivateProfileString(Section, KeyName, DefaultValue, Holder, 254, App.Path & "Config.ini")
    MyGetSetting = Left$(Holder, InStr(Holder, Chr$(0)) - 1)
Else
    X = GetPrivateProfileString(Section, KeyName, DefaultValue, Holder, 254, App.Path & "\Config.ini")
    MyGetSetting = Left$(Holder, InStr(Holder, Chr$(0)) - 1)
End If
End Function

'保存ini
'MySetSetting "项", "名称", Option1.Value
Public Sub MySetSetting(Section As String, KeyName As String, KeyValue As String)
Dim X As Long
If Right(App.Path, 2) = ":\" Then
    X = WritePrivateProfileString(Section, KeyName, KeyValue, App.Path & "Config.ini")
Else
    X = WritePrivateProfileString(Section, KeyName, KeyValue, App.Path & "\Config.ini")
End If
End Sub
