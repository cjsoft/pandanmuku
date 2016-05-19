VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command9 
      Caption         =   "1"
      Height          =   255
      Left            =   480
      TabIndex        =   8
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command5 
      Caption         =   "V"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command8 
      Caption         =   "2"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   480
      Width           =   255
   End
   Begin VB.CommandButton Command4 
      Caption         =   ">"
      Height          =   255
      Left            =   480
      TabIndex        =   3
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton Command7 
      Caption         =   "1"
      Height          =   255
      Left            =   480
      TabIndex        =   6
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command1 
      Caption         =   "+"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton Command3 
      Caption         =   "<"
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   240
      Width           =   255
   End
   Begin VB.CommandButton Command2 
      Caption         =   "A"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.CommandButton Command6 
      Caption         =   "2"
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call SaveSetting("CJSDMS", "FORM", "TOP", Str(Me.Top))
    Call SaveSetting("CJSDMS", "FORM", "LEFT", Str(Me.Left))
    Call SaveSetting("CJSDMS", "FORM", "WIDTH", Str(Me.Width))
    Call SaveSetting("CJSDMS", "FORM", "HEIGHT", Str(Me.Height))
End Sub

Private Sub Command2_Click()
    Me.Top = 0
    'Me.Width = Screen.Width
    
End Sub

Private Sub Command3_Click()
Me.Left = 0
'Me.Width = Screen.Width
End Sub

Private Sub Command4_Click()
Me.Left = Screen.Width - Me.Width
End Sub

Private Sub Command5_Click()
Me.Top = Screen.Height - Me.Height
End Sub

Private Sub Command6_Click()
Me.Width = Screen.Width / 2
End Sub

Private Sub Command7_Click()
Me.Width = Screen.Width
End Sub

Private Sub Command8_Click()
Me.Height = Screen.Height / 2
End Sub

Private Sub Command9_Click()
Me.Height = Screen.Height
End Sub

Private Sub Form_Load()
Form1.Top = CLng(GetSetting("CJSDMS", "FORM", "TOP", "0"))
    Form1.Left = CLng(GetSetting("CJSDMS", "FORM", "LEFT", "0"))
    Form1.Width = CLng(GetSetting("CJSDMS", "FORM", "WIDTH", Str(Screen.Width)))
    Form1.Height = CLng(GetSetting("CJSDMS", "FORM", "HEIGHT", Str(Screen.Height)))
End Sub
