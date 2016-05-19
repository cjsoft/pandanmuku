Attribute VB_Name = "mdlCen"
Dim danmu As New Collection
Dim Iter As Long
Type Slot
    Top As Integer
    Cap As Boolean
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
    If Form1.WindowState = 1 Then Exit Sub
    Dim Height As Integer
    Form1.Top = CLng(GetSetting("CJSDMS", "FORM", "TOP", "0"))
    Form1.Left = CLng(GetSetting("CJSDMS", "FORM", "LEFT", "0"))
    Form1.Width = CLng(GetSetting("CJSDMS", "FORM", "WIDTH", str(Screen.Width)))
    Form1.Height = CLng(GetSetting("CJSDMS", "FORM", "HEIGHT", str(Screen.Height)))
    Height = Form1.ScaleHeight - Form1.ScaleHeight Mod Form1.lModel.Height
    ReDim Slots(IIf(Height / Form1.lModel.Height - 1 = -1, 0, Height / Form1.lModel.Height - 1))
    For i = LBound(Slots) To UBound(Slots)
        Slots(i).Top = i * Form1.lModel.Height
        Slots(i).Cap = False
    Next
End Sub

Sub Clear()
    On Error Resume Next
    For Each labl In danmu
        Controls.Remove labl
    Next
    Set danmu = Nothing
    Set danmu = New Collection
End Sub

Sub MessageMove(ByVal dX As Integer)
    On Error Resume Next
    Dim ODelete As New Collection
    For i = LBound(Slots) To UBound(Slots)
        Slots(i).Cap = False
    Next
    For Each a In danmu
        a.Left = a.Left - dX
        If a.Left + a.Width > Form1.ScaleWidth Then Slots(CInt(a.Tag)).Cap = True
        If a.Left = 0 - dX - a.Width Then Form1.Controls.Remove a
    Next
End Sub

Sub PopMessage(content As String, Icolor As Integer)
    On Error Resume Next
    Iter = Iter + 1
    Dim Temp As String
    Dim ObjName As String
    ObjName = "danmu" & Iter & "at" & GetStamp()
    Temp = Iter
    Temp = Replace(Temp, "-", "_")
    Dim tobj As Object
    Set tobj = Form1.Controls.Add("VB.Label", ObjName, Form1)
    tobj.Caption = content
    tobj.AutoSize = True
    tobj.Left = Form1.ScaleWidth
    tobj.Tag = "" '
    For i = 0 To UBound(Slots)
        If Slots(i).Cap = False Then
            tobj.Tag = str(i)
            Exit For
        End If
    Next
    If tobj.Tag = "" Then tobj.Tag = str(Random(UBound(Slots)))
    tobj.Top = Slots(CInt(tobj.Tag)).Top
    tobj.Font = Form1.lModel.Font
    tobj.Fontsize = Form1.lModel.Fontsize
    tobj.BackStyle = 0
    Slots(CInt(tobj.Tag)).Cap = True
    tobj.ForeColor = Colo(Random(UBound(Colo) + 1)) 'IIf(Icolor = -1, Val(Colo(Random(UBound(Colo)))), Icolor)
    'tobj.Tag = Random(25)
    danmu.Add Form1.Controls(ObjName)
    tobj.Visible = True
End Sub

Sub CheckMailAndPop()
  
End Sub



