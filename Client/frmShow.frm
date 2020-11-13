VERSION 5.00
Begin VB.Form frmShow 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   435
      Left            =   3090
      TabIndex        =   4
      Top             =   2310
      Width           =   1245
   End
   Begin VB.TextBox txtPassword 
      Height          =   315
      Left            =   2070
      TabIndex        =   1
      Top             =   1080
      Width           =   2265
   End
   Begin VB.TextBox txtUserName 
      Height          =   315
      Left            =   2070
      TabIndex        =   0
      Top             =   630
      Width           =   2265
   End
   Begin VB.Label Label2 
      Caption         =   "Password"
      Height          =   255
      Left            =   570
      TabIndex        =   3
      Top             =   1110
      Width           =   1305
   End
   Begin VB.Label Label1 
      Caption         =   "User Name"
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Top             =   660
      Width           =   1365
   End
End
Attribute VB_Name = "frmShow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim UserMaintain As clsUserMainain

Private Sub cmdSave_Click()
    Set UserMaintain = New clsUserMainain
    UserMaintain.User_Name = txtUserName.Text
    UserMaintain.Pass_Ward = txtPassword.Text
    Call UserMaintain.Pass_UserMaintain(UserMaintain.User_Name, UserMaintain.Pass_Ward)
    txtUserName.Text = ""
    txtPassword.Text = ""
    txtUserName.SetFocus
    MsgBox "Record Saved.", vbInformation + vbOKOnly, "Record Saved"
End Sub

Private Sub Form_Activate()
    txtUserName.SetFocus
End Sub

