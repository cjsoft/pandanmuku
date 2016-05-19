VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "CJSoft Danmaku System By 褚尼兔"
   ClientHeight    =   5475
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9015
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5475
   ScaleWidth      =   9015
   Begin VB.Timer tmrDanmUKongzhi 
      Interval        =   42
      Left            =   8040
      Top             =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   8520
      Top             =   0
   End
   Begin VB.Label lModel 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   36
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   6120
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim acl()

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private Const HWND_TOPMOST& = -1
' 将窗口置于列表顶部，并位于任何最顶部窗口的前面
Private Const SWP_NOSIZE& = &H1
' 保持窗口大小
Private Const SWP_NOMOVE& = &H2
' 保持窗口位置
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private Const WS_EX_LAYERED = &H80000
Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const LWA_COLORKEY = &H1

Dim WithEvents chk As PandanmukuCilent
Attribute chk.VB_VarHelpID = -1



Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        chk.quit
        Clear
        Set chk = Nothing
        End
    ElseIf KeyAscii = 13 Then
        
    End If
End Sub

Private Sub Form_Load()
    If App.PrevInstance Then End
    
    ReadSettings
    Me.Top = CInt(GetSetting("CJSDMS", "FORM", "TOP", "0"))
    Me.Left = CInt(GetSetting("CJSDMS", "FORM", "LEFT", "0"))
    Me.Width = CInt(GetSetting("CJSDMS", "FORM", "WIDTH", str(Screen.Width)))
    Me.Height = CInt(GetSetting("CJSDMS", "FORM", "HEIGHT", str(Screen.Height)))
    BackTransparent
    Set chk = New PandanmukuCilent
    chk.SetServer se, pt, lpt
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrDanmUKongzhi.Enabled = False
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    tmrDanmUKongzhi.Enabled = True
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        Clear
        chk.quit
        Set chk = Nothing
        End
End Sub

Private Sub Form_Resize()
If Me.WindowState <> 1 Then Resize
End Sub

Private Sub Timer1_Timer()
    Dim ast As String, bst As String, cst As Boolean
    ast = unn: bst = pww
    ast = GetSetting("CJSDMS", "SERVER", "USERNAME", "")
    bst = GetSetting("CJSDMS", "SERVER", "PASSWORD", "")
    
    If (ast = "" Or bst = "") Then
        cst = chk.RegisternLogin(ast, bst)
        InputBox "", , ast
        Call SaveSetting("CJSDMS", "SERVER", "USERNAME", ast)
        Call SaveSetting("CJSDMS", "SERVER", "PASSWORD", bst)
    Else
        cst = chk.RegisternLogin(ast, bst)
    End If
    Timer1.Enabled = False
End Sub

Private Sub tmrDanmUKongzhi_Timer()
    MessageMove SingleMove
End Sub

Private Sub chk_PostPDM(content, col As Integer)
        If col > UBound(Colo) Or col < -1 Then col = Random(UBound(Colo))
        Dim stemp As String
        stemp = content
        PopMessage stemp, col

End Sub


Sub BackTransparent()
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
' 将窗口设为在所有窗口前端
  Me.BackColor = &H0
   Dim rtn As Long
   Dim BorderStyler
   BorderStyler = 0
   rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
   rtn = rtn Or WS_EX_LAYERED
   SetWindowLong hwnd, GWL_EXSTYLE, rtn
   SetLayeredWindowAttributes hwnd, &H0, 0, LWA_COLORKEY
End Sub

