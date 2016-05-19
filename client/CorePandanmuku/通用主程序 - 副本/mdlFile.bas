Attribute VB_Name = "mdlFile"
Public Colo() As String
Public se As String, us As String, pw As String, pt As Long, lpt As Long
Dim Fontsize As Integer
Dim Font
Dim Colors()
Public SingleMove As Integer
Sub ReadSettings()
    On Error Resume Next
    Font = R("FONT", "NAME")
    Fontsize = Val(R("FONT", "SIZE"))
    Colo() = Split(R("FONT", "COLORLIST"), ",", , vbTextCompare)
    SingleMove = Val(R("CONTROL", "SINGLEMOVE"))
    Form1.Timer1.Interval = Val(R("CONTROL", "REFRESHINTERVAL"))
    Form1.lModel.FontName = Font
    Form1.lModel.Fontsize = Fontsize
    se = R("SERVER", "ADDRESS")
    pt = CLng(R("SERVER", "PORT"))
    lpt = CLng(R("SERVER", "LPORT"))
    us = R("SERVER", "USERNAME")
    pw = R("SERVER", "PASSWORD")
End Sub
Function GetKeyValueFromCJSX(ByVal FilePath As String, ByVal SectionName As String, ByVal KeyName As String)
    Dim FileContent As String
    Dim FileContentPreloader
    Dim DataStorage() As String
    Dim SectionNa As String
    Dim KeyNa As String
    Dim EQualPos As Integer
    GetKeyValueFromCJSX = "#No Such a Key Found#"

    If Dir(FilePath, vbNormal) <> "" Or Dir(FilePath, vbHidden) <> "" Or Dir(FilePath, vbReadOnly) <> "" Then
        Open FilePath For Input As #1
        Debug.Print "已找到指定文件，正整理文件内容"
        Do While Not EOF(1)
            Line Input #1, FileContentPreloader
            FileContent = FileContent & vbCrLf & FileContentPreloader
        Loop
        Close #1
        DataStorage() = Split(FileContent, vbCrLf, , vbTextCompare)
        Debug.Print "文件内容整理完毕，正寻找目标Section"
        
        'TestStarted-----------------------------
        For i = LBound(DataStorage) To UBound(DataStorage)
            If Trim(DataStorage(i)) = "" Then GoTo InstantNext
                If Left(Trim(DataStorage(i)), 1) = "<" And Right(Trim(DataStorage(i)), 1) = ">" And Len(Trim(DataStorage(i)) >= 3) Then
                    SectionNa = Mid(Trim(DataStorage(i)), 2, Len(Trim(DataStorage(i))) - 2)
                    GoTo InstantNext
                End If
                If SectionNa = SectionName Then
                    Debug.Print "已找到目标Section """ & SectionNa & """，正在寻找目标Key"
                    If Len(Trim(DataStorage(i))) >= 2 Then
                        EQualPos = InStr(2, Trim(DataStorage(i)), "=", vbTextCompare)
                        If EQualPos <> 0 Then
                            KeyNa = Mid(Trim(DataStorage(i)), 1, EQualPos - 1)
                            Debug.Print "KeyNa:""" & KeyNa & """"
                            If KeyNa = KeyName Then
                                If Len(Trim(DataStorage(i))) = equanlpos Then
                                    GetKeyValueFromCJSX = ""
                                Else
                                    GetKeyValueFromCJSX = Trim(Right(Trim(DataStorage(i)), Len(Trim(DataStorage(i))) - EQualPos))
                                    
                                End If
                                Debug.Print "最终结果：""" & GetKeyValueFromCJSX & """"
                                Exit Function
                            End If
                        End If
                    End If
                End If
InstantNext:
        Next
        'TestEnded-------------------------------
        
        
        'Some code for the file-reading function
    Else
        GetKeyValueFromCJSX = "#No Such a CJSX File Found#"
    End If
    If GetKeyValueFromCJSX = "#No Such a Key Found#" Then Debug.Print "没有找到目标Key"
End Function
Function PutKeyValueIntoCJSX(ByVal FilePath As String, ByVal SectionName As String, ByVal KeyName As String, ByVal KeyValue As String)
    Dim FileContent As String
    Dim FileContentPreloader
    Dim DataStorage() As String
    Dim SectionNa As String
    Dim KeyNa As String
    Dim EQualPos As Integer
    PutKeyValueIntoCJSX = "#No Such a Key Found#"
Redo:
    If Dir(FilePath, vbNormal) <> "" Or Dir(FilePath, vbHidden) <> "" Or Dir(FilePath, vbReadOnly) <> "" Then
        Open FilePath For Input As #1
        Debug.Print "已找到指定文件，正整理文件内容"
        Do While Not EOF(1)
            Line Input #1, FileContentPreloader
            FileContent = FileContent & vbCrLf & FileContentPreloader
        Loop
        Close #1
        DataStorage() = Split(FileContent, vbCrLf, , vbTextCompare)
        Debug.Print "文件内容整理完毕，正寻找目标Section"
        
        'TestStarted-----------------------------
        For i = LBound(DataStorage) To UBound(DataStorage)
            If Trim(DataStorage(i)) = "" Then GoTo InstantNext
                If Left(Trim(DataStorage(i)), 1) = "<" And Right(Trim(DataStorage(i)), 1) = ">" And Len(Trim(DataStorage(i)) >= 3) Then
                    If SectionNa = SectionName Then
                    
                        SectionNa = Mid(Trim(DataStorage(i)), 2, Len(Trim(DataStorage(i))) - 2)
                        If SectionNa <> SectionName And PutKeyValueIntoCJSX = "#No Such a Key Found#" Then
                            DataStorage(i) = KeyName & "=" & KeyValue & vbCrLf & "<" & SectionNa & ">"
                            Open FilePath For Output As #1
                            For o = LBound(DataStorage) To UBound(DataStorage)
                                If Trim(DataStorage(o)) <> "" Then
                                    Print #1, DataStorage(o)
                                End If
                            Next
                            Close #1
                            PutKeyValueIntoCJSX = "#Appended#"
                            Exit Function
                        End If
                    Else
                        SectionNa = Mid(Trim(DataStorage(i)), 2, Len(Trim(DataStorage(i))) - 2)
                        GoTo InstantNext
                    End If
                End If
                If SectionNa = SectionName Then
                    Debug.Print "已找到目标Section """ & SectionNa & """，正在寻找目标Key"
                    If Len(Trim(DataStorage(i))) >= 2 Then
                        EQualPos = InStr(2, Trim(DataStorage(i)), "=", vbTextCompare)
                        If EQualPos <> 0 Then
                            KeyNa = Mid(Trim(DataStorage(i)), 1, EQualPos - 1)
                            Debug.Print "KeyNa:""" & KeyNa & """"
                            If KeyNa = KeyName Then
                                DataStorage(i) = KeyName & "=" & KeyValue
                                
                                    Open FilePath For Output As #1
                                    For o = LBound(DataStorage) To UBound(DataStorage)
                                        If Trim(DataStorage(o)) <> "" Then
                                            Print #1, DataStorage(o)
                                        End If
                                    Next
                                    Close #1
                                    PutKeyValueIntoCJSX = "#Changed#"
                                Exit Function
                            End If
                        End If
                    End If
                
                    If PutKeyValueIntoCJSX = "#No Such a Key Found#" Then
                        Open FilePath For Output As #1
                        For o = LBound(DataStorage) To UBound(DataStorage)
                           If Trim(DataStorage(o)) <> "" Then
                                Print #1, DataStorage(o)
                            End If
                        Next
                        Print #1, KeyName & "=" & KeyValue
                        Close #1
                        PutKeyValueIntoCJSX = "#Changed#"
                        Exit Function
                    End If
                End If
InstantNext:
        Next
        'TestEnded-------------------------------
        If PutKeyValueIntoCJSX = "#No Such a Key Found#" Then
            Open FilePath For Output As #1
            For o = LBound(DataStorage) To UBound(DataStorage)
                If Trim(DataStorage(o)) <> "" Then
                    Print #1, DataStorage(o)
                End If
            Next
            Print #1, "<" & SectionName & ">"
            Print #1, KeyName & "=" & KeyValue
            Close #1
            PutKeyValueIntoCJSX = "#Appended#"
        End If
        
        'Some code for the file-reading function
    Else
        Open FilePath For Output As #1
        
        Close #1
        GoTo Redo
    End If
    If PutKeyValueIntoCJSX = "#No Such a Key Found#" Then Debug.Print "没有找到目标Key"

End Function


Public Function R(S As String, K As String)
    R = GetKeyValueFromCJSX(Replace(App.Path & "\settings.cjsx", "\\", "\"), S, K)
End Function
Public Function P(S As String, K As String, V As String)
    P = PutKeyValueIntoCJSX(Replace(App.Path & "\settings.cjsx", "\\", "\"), S, K, V)
End Function
