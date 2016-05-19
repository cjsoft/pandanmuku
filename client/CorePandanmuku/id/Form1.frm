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
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Call SaveSetting("CJSDMS", "SERVER", "USERNAME", InputBox("username:", "PDMS id查看修改器", GetSetting("CJSDMS", "SERVER", "USERNAME")))
    Call SaveSetting("CJSDMS", "SERVER", "PASSWORD", InputBox("password:", "PDMS id查看修改器", GetSetting("CJSDMS", "SERVER", "PASSWORD")))
    End
End Sub
