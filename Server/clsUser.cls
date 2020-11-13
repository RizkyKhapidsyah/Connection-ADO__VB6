VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Private disp_UserName As String
Private disp_PassWord As String
Private disp_aaa As String

Private mvarUserName As String
Private mvarPassWord As String
Private mvaraaa As String 'local copy
Public Property Let aaa(ByVal vData As String)
    mvaraaa = vData
End Property

Public Property Get aaa() As String
    aaa = mvaraaa
End Property

Public Function FbooConnection() As Boolean
Dim strConnection As String

On Error GoTo Err_FbooConnection

 Set gobjConnection = New ADODB.Connection
'' strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Connection\Database\Testing.mdb;Persist Security Info=False"
 
 strConnection = "Provider=MSDAORA.1;User ID='" & mvarUserName & "';PASSWORD='" & mvarPassWord & "';Data Source=TEST;Persist Security Info=False"
 gobjConnection.Open strConnection
 FbooConnection = True
 
 Exit Function
 
Err_FbooConnection:
    FbooConnection = False
End Function

Public Property Let PassWord(ByVal vData As String)
    mvarPassWord = vData
End Property

Public Property Get PassWord() As String
    PassWord = mvarPassWord
End Property

Public Property Let UserName(ByVal vData As String)
    mvarUserName = vData
End Property

Public Property Get UserName() As String
    UserName = mvarUserName
End Property

Public Function FLoginSuccess() As Boolean
 
 On Error GoTo Err_FLoginSuccess
    
    Set rstRecordSet = New ADODB.Recordset
    strRecordSetString = "Select * from dpmc.tblTTest"
    rstRecordSet.Open strRecordSetString, gobjConnection
    
    If rstRecordSet.EOF = False Then
        disp_UserName = rstRecordSet.Fields(0).Value
        disp_PassWord = rstRecordSet.Fields(1).Value
        FLoginSuccess = True
    Else
        MsgBox "No records found.", vbInformation + vbOKOnly, "No Records"
    End If
    
    Exit Function

Err_FLoginSuccess:
    FLoginSuccess = False
End Function


Public Function FDispayInformation(m_UserName, m_PassWord)
    m_UserName = disp_UserName
    m_PassWord = disp_PassWord
End Function

Public Function Add() As Boolean

On Error GoTo Err_Add
    
    Set rstRecordSet = New ADODB.Recordset
    strRecordSetString = "Insert into admUserAdmin values ('" & mvarUserName & "', '" & mvarPassWord & "','" & mvaraaa & "')"
    rstRecordSet.Open strRecordSetString, gobjConnection
    Add = True
    Exit Function
    
Err_Add:
    MsgBox Err.Description
    Add = False
End Function

Public Function Delete() As Boolean

On Error GoTo Err_Delete
    
    Set rstRecordSet = New ADODB.Recordset
    strRecordSetString = "Delete from admUserAdmin where UserName = '" & mvarUserName & "'"
    rstRecordSet.Open strRecordSetString, gobjConnection
    Delete = True
    Exit Function
    
Err_Delete:
    Delete = False
End Function

