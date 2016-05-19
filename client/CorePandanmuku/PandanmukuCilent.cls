VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PandanmukuCilent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private mfrm As frmmain
Private WithEvents T As Timer
Attribute T.VB_VarHelpID = -1
Private WithEvents wsk As Winsock
Attribute wsk.VB_VarHelpID = -1
Public Event PostPDM(content, row As Integer)
Dim unn As String, pww As String
Dim RTAG
Private Sub Class_Initialize()
    Set mfrm = New frmmain
    Set T = frmmain.Timer1
    Set wsk = frmmain.wsk
    Load mfrm
    Set OriJSON = New Dictionary
    RTAG = True
    ifLoginned = False
End Sub


Private Sub Class_Terminate()
    Set T = Nothing
    End
End Sub


Public Sub SetServer(ByVal HostAddress As String, ByVal HostPort As Long, ByVal LocalPort As Long)
    wsk.LocalPort = LocalPort
    wsk.RemoteHost = HostAddress
    wsk.RemotePort = HostPort
End Sub

Public Function RegisternLogin(ByRef UserName, ByRef PassWord) As Boolean
    Dim templ As Long
    templ = 0
    On Error GoTo ERRT
    Dim oj As Object
    Set oj = OriJSON
    With oj
        
        unn = UserName
        pww = PassWord
        If unn <> "" And pww <> "" Then GoTo LOGG
        .Item("command") = "register"

        wsk.SendData (JSON.toString(oj))
        DoEvents
        Do While RTAG
            templ = templ + 1
            If (templ > 5000000) Then
                MsgBox "超时，服务器可能没有运行，请联系中心服务器管理员", vbCritical
                End
                Exit Do
            End If
            DoEvents
            UserName = unn
            PassWord = pww
        Loop
LOGG:
        
        .Item("command") = "login"
        .Item("username") = unn
        .Item("password") = pww
        wsk.SendData (JSON.toString(oj))
        templ = 0
        Do While RTAG
            templ = templ + 1
            If (templ > 5000000) Then
                MsgBox "超时，服务器可能没有运行，请联系中心服务器管理员", vbCritical
                End
            End If
            DoEvents
        Loop
    End With
    RegisternLogin = True
    T.Enabled = True
    Exit Function
ERRT:
    RegisternLogin = False
End Function
Private Sub wsk_DataArrival(ByVal bytesTotal As Long)
    
    On Error GoTo ERRT
    Dim DataBuffer() As Byte
    wsk.GetData DataBuffer
    Dim TSTR As String
    TSTR = Bytes_to_Unicode(DataBuffer, "UTF-8")
    Open "logging.log" For Append As #1
    Debug.Print TSTR
    Print #1, TSTR
    Close #1
    Dim oj As Object
    Set oj = JSON.parse(TSTR)
    If Not (False) Then
        If oj.Item("type") = "ok" Then
            Dim ta As String, tb As String
            ta = oj.Item("tag")
            If InStr(1, ta, "successfully registered") <> 0 Then
                unn = oj.Item("username")
                pww = oj.Item("password")
                RTAG = False
                Exit Sub
            End If
            If InStr(1, oj.Item("tag"), "successfully loginned") <> 0 Then
                ifLoginned = True
                RTAG = False
                Exit Sub
            End If
        ElseIf oj.Item("type") = "pandanmuku" Then
            RaiseEvent PostPDM(oj.Item("content"), -1)
        ElseIf oj.Item("type") = "error" Then
            If InStr(1, oj.Item("tag"), "login failed") <> 0 Then
                MsgBox "错误的登录用户名或密码，请执行id信息清除（如果服务器开放注册），或者联系中心服务器管理员进行手动注册" & vbCrLf & "您现在可以关闭本程序", vbCritical
                End
                Exit Sub
            End If
        End If
    End If
    Exit Sub
ERRT:
    
End Sub

Public Sub quit()
    Dim oj As Object
    Set oj = OriJSON
    oj.Item("command") = "logout"
    wsk.SendData (JSON.toString(oj))
End Sub
Private Sub T_Timer()
    Dim oj As Object
    Set oj = OriJSON
    oj.Item("command") = "heartbeat"
    wsk.SendData (JSON.toString(oj))
End Sub
