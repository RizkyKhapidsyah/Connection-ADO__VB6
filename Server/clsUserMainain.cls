VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUserMainain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private mvarUser_Name As String
Private mvarPass_Ward As String
Public Property Let Pass_Ward(ByVal vData As String)
    mvarPass_Ward = vData
End Property

Public Property Get Pass_Ward() As String
    Pass_Ward = mvarPass_Ward
End Property

Public Property Let User_Name(ByVal vData As String)
    mvarUser_Name = vData
End Property

Public Property Get User_Name() As String
    User_Name = mvarUser_Name
End Property

Public Function Pass_UserMaintain(UserName, PassWord)
 Dim cmd As ADODB.Command
 Dim param As ADODB.Parameter
 Dim PassWord1 As ADODB.Parameter
 Set cmd = New ADODB.Command
 
 Set cmd.ActiveConnection = gobjConnection
 cmd.CommandText = "UserMaintain"
 cmd.CommandType = adCmdStoredProc
 
 Set param = cmd.CreateParameter(UserName, adVarChar, adParamInput, 20, UserName)
 Set PassWord1 = cmd.CreateParameter(PassWord, adVarChar, adParamInput, 20, PassWord)
 
 cmd.Parameters.Append param
 cmd.Parameters.Append PassWord1
 
 cmd.Execute
 
End Function

