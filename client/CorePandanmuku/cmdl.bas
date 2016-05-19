Attribute VB_Name = "cmdl"
Public Pw As String, Un As String
Public OriJSON As Object
Public RTAG As Boolean
Public ifLoginned As Boolean

'字节数值转汉字
Public Function Bytes_to_Unicode(Bytes, CodeType As String)
    Dim strReturn As String
    Dim i As Long
    Dim ThisCharCode As Integer
    Dim NextCharCode As Integer
    Dim ThirdCharCode As Integer
    strReturn = ""
    For i = 1 To LenB(Bytes)
        ThisCharCode = AscB(MidB(Bytes, i, 1))
        If ThisCharCode < &H80 Then
            strReturn = strReturn & Chr(ThisCharCode)
        Else
            If CodeType = "UTF-8" Or CodeType = "UTF8" Then
                NextCharCode = AscB(MidB(Bytes, i + 1, 1))
                ThirdCharCode = AscB(MidB(Bytes, i + 2, 1))
                strReturn = strReturn & UTF8_to_Unicode(ThisCharCode, NextCharCode, ThirdCharCode)
                i = i + 2
            Else
                NextCharCode = AscB(MidB(Bytes, i + 1, 1))
                strReturn = strReturn & Unicode(ThisCharCode, NextCharCode)
                i = i + 1
            End If
        End If
    Next
    Bytes_to_Unicode = strReturn
End Function
'二字节汉字转换
Public Function Unicode(BY1, BY2) As String
    Unicode = Chr(Int(BY1) * 256 + Int(BY2))
End Function
'三字节的UTF-8编码转二字节的Unicode编码
Function UTF8_to_Unicode(BY1, BY2, BY3) As String
    Dim BIN_UTF8 As String
    BIN_UTF8 = DEC_to_BIN(Int(BY1)) & DEC_to_BIN(Int(BY2)) & DEC_to_BIN(Int(BY3))
    Dim BIN_Unicode As String
    BIN_Unicode = Mid(BIN_UTF8, 5, 4) & Mid(BIN_UTF8, 11, 6) & Mid(BIN_UTF8, 19, 6)
    Dim DEC_Unicode As Long
    DEC_Unicode = BIN_to_DEC(BIN_Unicode)
    UTF8_to_Unicode = ChrW(DEC_Unicode)
End Function
 
'双字节的转换，直接用第一个字节*256+第二个字节
'UTF-B的三字节二进制:1110xxxx,10xxxxxx,10xxxxxx
'Unicde的二字节对应的二进制是上面的x组合:xxxx+xxxxxx+xxxxxx
'将组合后的二进制转为10进制
'得到汉字



Public Function DEC_to_BIN(ByVal Dec As Long) As String
    DEC_to_BIN = ""
    Do While Dec > 0
        DEC_to_BIN = Dec Mod 2 & DEC_to_BIN
        Dec = Dec \ 2
    Loop
End Function

' 用途：将二 进 制转化为十进制
' 输入：Bin(二 进 制数)
' 输入数据类型：String
' 输出：BIN_to_DEC(十进制数)
' 输出数据类型：Long
' 输入的最大数为1111111111111111111111111111111(31个1),输出最大数为2147483647
Public Function BIN_to_DEC(ByVal Bin As String) As Long
    Dim i As Long
    For i = 1 To Len(Bin)
        BIN_to_DEC = BIN_to_DEC * 2 + Val(Mid(Bin, i, 1))
    Next i
End Function

' 用途：将十六进制转化为二 进 制
' 输入：Hex(十六进制数)
' 输入数据类型：String
' 输出：HEX_to_BIN(二 进 制数)
' 输出数据类型：String
' 输入的最大数为2147483647个字符
Public Function HEX_to_BIN(ByVal Hex As String) As String
    Dim i As Long
    Dim B As String
    
    Hex = UCase(Hex)
    For i = 1 To Len(Hex)
        Select Case Mid(Hex, i, 1)
            Case "0": B = B & "0000"
            Case "1": B = B & "0001"
            Case "2": B = B & "0010"
            Case "3": B = B & "0011"
            Case "4": B = B & "0100"
            Case "5": B = B & "0101"
            Case "6": B = B & "0110"
            Case "7": B = B & "0111"
            Case "8": B = B & "1000"
            Case "9": B = B & "1001"
            Case "A": B = B & "1010"
            Case "B": B = B & "1011"
            Case "C": B = B & "1100"
            Case "D": B = B & "1101"
            Case "E": B = B & "1110"
            Case "F": B = B & "1111"
        End Select
    Next i
    While Left(B, 1) = "0"
        B = Right(B, Len(B) - 1)
    Wend
    HEX_to_BIN = B
End Function

' 用途：将二 进 制转化为十六进制
' 输入：Bin(二 进 制数)
' 输入数据类型：String
' 输出：BIN_to_HEX(十六进制数)
' 输出数据类型：String
' 输入的最大数为2147483647个字符
Public Function BIN_to_HEX(ByVal Bin As String) As String
    Dim i As Long
    Dim H As String
    If Len(Bin) Mod 4 <> 0 Then
        Bin = String(4 - Len(Bin) Mod 4, "0") & Bin
    End If
    
    For i = 1 To Len(Bin) Step 4
        Select Case Mid(Bin, i, 4)
            Case "0000": H = H & "0"
            Case "0001": H = H & "1"
            Case "0010": H = H & "2"
            Case "0011": H = H & "3"
            Case "0100": H = H & "4"
            Case "0101": H = H & "5"
            Case "0110": H = H & "6"
            Case "0111": H = H & "7"
            Case "1000": H = H & "8"
            Case "1001": H = H & "9"
            Case "1010": H = H & "A"
            Case "1011": H = H & "B"
            Case "1100": H = H & "C"
            Case "1101": H = H & "D"
            Case "1110": H = H & "E"
            Case "1111": H = H & "F"
        End Select
    Next i
    While Left(H, 1) = "0"
        H = Right(H, Len(H) - 1)
    Wend
    BIN_to_HEX = H
End Function

' 用途：将十六进制转化为十进制
' 输入：Hex(十六进制数)
' 输入数据类型：String
' 输出：HEX_to_DEC(十进制数)
' 输出数据类型：Long
' 输入的最大数为7FFFFFFF,输出的最大数为2147483647
Public Function HEX_to_DEC(ByVal Hex As String) As Long
    Dim i As Long
    Dim B As Long
    
    Hex = UCase(Hex)
    For i = 1 To Len(Hex)
        Select Case Mid(Hex, Len(Hex) - i + 1, 1)
            Case "0": B = B + 16 ^ (i - 1) * 0
            Case "1": B = B + 16 ^ (i - 1) * 1
            Case "2": B = B + 16 ^ (i - 1) * 2
            Case "3": B = B + 16 ^ (i - 1) * 3
            Case "4": B = B + 16 ^ (i - 1) * 4
            Case "5": B = B + 16 ^ (i - 1) * 5
            Case "6": B = B + 16 ^ (i - 1) * 6
            Case "7": B = B + 16 ^ (i - 1) * 7
            Case "8": B = B + 16 ^ (i - 1) * 8
            Case "9": B = B + 16 ^ (i - 1) * 9
            Case "A": B = B + 16 ^ (i - 1) * 10
            Case "B": B = B + 16 ^ (i - 1) * 11
            Case "C": B = B + 16 ^ (i - 1) * 12
            Case "D": B = B + 16 ^ (i - 1) * 13
            Case "E": B = B + 16 ^ (i - 1) * 14
            Case "F": B = B + 16 ^ (i - 1) * 15
        End Select
    Next i
    HEX_to_DEC = B
End Function
' 用途：将十进制转化为十六进制
' 输入：Dec(十进制数)
' 输入数据类型：Long
' 输出：DEC_to_HEX(十六进制数)
' 输出数据类型：String
' 输入的最大数为2147483647,输出最大数为7FFFFFFF
Public Function DEC_to_HEX(Dec As Long) As String
    Dim a As String
    DEC_to_HEX = ""
    Do While Dec > 0
        a = CStr(Dec Mod 16)
        Select Case a
            Case "10": a = "A"
            Case "11": a = "B"
            Case "12": a = "C"
            Case "13": a = "D"
            Case "14": a = "E"
            Case "15": a = "F"
        End Select
        DEC_to_HEX = a & DEC_to_HEX
        Dec = Dec \ 16
    Loop
End Function

' 用途：将十进制转化为八进制
' 输入：Dec(十进制数)
' 输入数据类型：Long
' 输出：DEC_to_OCT(八进制数)
' 输出数据类型：String
' 输入的最大数为2147483647,输出最大数为17777777777
Public Function DEC_to_OCT(ByVal Dec As Long) As String
    DEC_to_OCT = ""
    Do While Dec > 0
        DEC_to_OCT = Dec Mod 8 & DEC_to_OCT
        Dec = Dec \ 8
    Loop
End Function

' 用途：将八进制转化为十进制
' 输入：Oct(八进制数)
' 输入数据类型：String
' 输出：OCT_to_DEC(十进制数)
' 输出数据类型：Long
' 输入的最大数为17777777777,输出的最大数为2147483647
Public Function OCT_to_DEC(ByVal Oct As String) As Long
    Dim i As Long
    Dim B As Long
    
    For i = 1 To Len(Oct)
        Select Case Mid(Oct, Len(Oct) - i + 1, 1)
            Case "0": B = B + 8 ^ (i - 1) * 0
            Case "1": B = B + 8 ^ (i - 1) * 1
            Case "2": B = B + 8 ^ (i - 1) * 2
            Case "3": B = B + 8 ^ (i - 1) * 3
            Case "4": B = B + 8 ^ (i - 1) * 4
            Case "5": B = B + 8 ^ (i - 1) * 5
            Case "6": B = B + 8 ^ (i - 1) * 6
            Case "7": B = B + 8 ^ (i - 1) * 7
        End Select
    Next i
    OCT_to_DEC = B
End Function

' 用途：将二 进 制转化为八进制
' 输入：Bin(二 进 制数)
' 输入数据类型：String
' 输出：BIN_to_OCT(八进制数)
' 输出数据类型：String
' 输入的最大数为2147483647个字符
Public Function BIN_to_OCT(ByVal Bin As String) As String
    Dim i As Long
    Dim H As String
    If Len(Bin) Mod 3 <> 0 Then
        Bin = String(3 - Len(Bin) Mod 3, "0") & Bin
    End If
    
    For i = 1 To Len(Bin) Step 3
        Select Case Mid(Bin, i, 3)
            Case "000": H = H & "0"
            Case "001": H = H & "1"
            Case "010": H = H & "2"
            Case "011": H = H & "3"
            Case "100": H = H & "4"
            Case "101": H = H & "5"
            Case "110": H = H & "6"
            Case "111": H = H & "7"
        End Select
    Next i
    While Left(H, 1) = "0"
        H = Right(H, Len(H) - 1)
    Wend
    BIN_to_OCT = H
End Function

' 用途：将八进制转化为二 进 制
' 输入：Oct(八进制数)
' 输入数据类型：String
' 输出：OCT_to_BIN(二 进 制数)
' 输出数据类型：String
' 输入的最大数为2147483647个字符
Public Function OCT_to_BIN(ByVal Oct As String) As String
    Dim i As Long
    Dim B As String
    
    For i = 1 To Len(Oct)
        Select Case Mid(Oct, i, 1)
            Case "0": B = B & "000"
            Case "1": B = B & "001"
            Case "2": B = B & "010"
            Case "3": B = B & "011"
            Case "4": B = B & "100"
            Case "5": B = B & "101"
            Case "6": B = B & "110"
            Case "7": B = B & "111"
        End Select
    Next i
    While Left(B, 1) = "0"
        B = Right(B, Len(B) - 1)
    Wend
    OCT_to_BIN = B
End Function

' 用途：将八进制转化为十六进制
' 输入：Oct(八进制数)
' 输入数据类型：String
' 输出：OCT_to_HEX(十六进制数)
' 输出数据类型：String
' 输入的最大数为2147483647个字符
Public Function OCT_to_HEX(ByVal Oct As String) As String
    Dim Bin As String
    Bin = OCT_to_BIN(Oct)
    OCT_to_HEX = BIN_to_HEX(Bin)
End Function

' 用途：将十六进制转化为八进制
' 输入：Hex(十六进制数)
' 输入数据类型：String
' 输出：HEX_to_OCT(八进制数)
' 输出数据类型：String
' 输入的最大数为2147483647个字符
Public Function HEX_to_OCT(ByVal Hex As String) As String
    Dim Bin As String
    Hex = UCase(Hex)
    Bin = HEX_to_BIN(Hex)
    HEX_to_OCT = BIN_to_OCT(Bin)
End Function
