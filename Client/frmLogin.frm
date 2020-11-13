VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "Form1"
   ClientHeight    =   3390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3390
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLogin 
      Caption         =   "&Login"
      Height          =   405
      Left            =   2040
      TabIndex        =   2
      Top             =   1830
      Width           =   1035
   End
   Begin VB.TextBox txtPassWord 
      Height          =   345
      Left            =   2070
      TabIndex        =   1
      Top             =   1200
      Width           =   2415
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   2070
      TabIndex        =   0
      Top             =   660
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   315
      Left            =   660
      TabIndex        =   4
      Top             =   1230
      Width           =   915
   End
   Begin VB.Label Label1 
      Caption         =   "UserName"
      Height          =   315
      Left            =   660
      TabIndex        =   3
      Top             =   690
      Width           =   1005
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim userLogin As clsUser

Private Sub cmdLogin_Click()
Dim strUser As String
Dim strPassWord As String

    Set userLogin = New clsUser
    
        userLogin.UserName = txtUserName.Text
        userLogin.Password = txtPassword.Text
        
    If userLogin.FbooConnection = True Then
   
        If userLogin.FLoginSuccess = True Then
''            Call userLogin.FDispayInformation(strUser, strPassWord)
            frmShow.Show
        End If
        
    Else
        MsgBox "Conenction fail"
    End If
    
    txtUserName = ""
    txtPassword = ""
    
End Sub

Private Sub cmdShow_Click()
    frmShow.Show
End Sub

