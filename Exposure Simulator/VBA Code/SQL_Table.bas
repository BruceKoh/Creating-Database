Attribute VB_Name = "SQL_Table"
Option Explicit
Private Conn As ADODB.Connection

Function ConnectToDB(Server As String, uName As String, pWord As String) As Boolean
 
    Set Conn = New ADODB.Connection
    On Error Resume Next
    
    Conn.ConnectionString = "Driver={SQL Server};Server=" & Server & ";" & _
    "Uid=" & uName & ";" & "Pwd=" & pWord
    Conn.Open
    
    If Conn.State = 0 Then
        ConnectToDB = False
    Else
        ConnectToDB = True
    End If
 
End Function

Function Query(SQL As String)
 
    Dim recordSet As ADODB.recordSet
    Dim Field As ADODB.Field

    Set recordSet = New ADODB.recordSet
    recordSet.Open SQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
    
End Function

Public Sub Create_Table()
 
    Dim SQL_Create_Table As String
    Dim Connected As Boolean
    Dim strUserName As String
    
    strUserName = Application.UserName
    strUserName = Replace(strUserName, " ", "_")
    'CREATING TABLE
    SQL_Create_Table = "CREATE TABLE TestCompare.dbo." & strUserName & "(Owner VARCHAR(8) NULL," & _
    "SKU VARCHAR(50) NULL,PartRev VARCHAR(10) NULL,Category VARCHAR(10) NULL," & _
    "Component VARCHAR(50) NULL,Description VARCHAR(100) NULL,Per_Rate INT NULL)"
        
    Connected = ConnectToDB(dbAddress, uName, pWord)
    
    If Connected Then
        Query SQL_Create_Table
        Conn.Close
    Else
        MsgBox "We have a problem!"
    End If
       
End Sub

