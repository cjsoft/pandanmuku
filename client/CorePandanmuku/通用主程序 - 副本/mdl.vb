Dim danmu As New Collection
Dim Iter As Long
Type Slot
    Top As Integer
End Type
Dim Slots() As Slot


Function GetStamp() As String
    GetStamp = DateDiff("s", "1970-01-01 00:00:00", Now)
End Function

Function Random(n As Long) As Long
    Randomize
    Random = Int(Rnd * n)
End Function


Sub Init()
    Resize
    Set danmu = New Collection
End Sub


Sub Resize()
    Dim Height As Integer
    Height = Form1.ScaleHeight - Form1.ScaleHeight Mod Form1.lModel.Height
    ReDim Slots(Height / Form1.lModel.Height - 1)
    For i = LBound(Slots) To UBound(Slots)
        Slots(i).Top = (i) * Form1.lModel.Height
    Next
End Sub

Sub Clear()
    For Each labl In danmu
        Controls.Remove labl
    Next
    Set danmu = Nothing
    Set danmu = New Collection
End Sub

Sub MessageMove(ByVal dX As Integer)
    Dim ODelete As New Collection
    For i = 1 To danmu.Count
        danmu.Item(i).Left = danmu.Item(i).Left - dX
        If danmu.Item(i).Left = 0 - dX - danmu.Item(i).Width Then ODelete.Add i
    Next
    For Each a In ODelete
        Controls.Remove danmu.Item(i)
        danmu.Remove i
    Next
End Sub

Sub PopMessage(Content As String)
    Dim Colo()
    Colo = Array(vbRed, vbBlue, vbWhite, vbpurple, vbCyan)
    Iter = Iter + 1
    Dim Temp As String
    Dim ObjName As String
    ObjName = "danmu" & Iter & "at" & GetStamp()
    Temp = Iter
    Temp = Replace(Temp, "-", "_")
    Dim tobj As Object
    Set tobj = Form1.Controls.Add("VB.Label", ObjName, Form1)
    tobj.Caption = Content
    tobj.Left = Form1.ScaleWidth
    tobj.Top = Slots(Random(UBound(Slots))).Top
    tobj.Font = Form1.lModel.Font
    tobj.Font = Form1.lModel.FontSize
    tobj.BackStyle = 0
    tobj.ForeColor = Colo(Random(5))
    danmu.Add Form1.Controls(ObjName)
    tobj.Visible = True
End Sub

Sub CheckMailAndPop()
    Dim mails As New jmail.pop3
    mails.Connect "e241410@163.com", "airport", "pop.163.com"
    If mails.Count > 0 Then
        For i = 1 To mails.Count
            PopMessage mails.Messages.Item(i).Body
        Next
        mails.deletemessages
    End If
    mails.Disconnect
End Sub

