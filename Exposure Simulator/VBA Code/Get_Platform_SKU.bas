Attribute VB_Name = "Get_Platform_SKU"
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

Function Query(SQL As String, col As Long)
 
    Dim recordSet As ADODB.recordSet
    Dim Field As ADODB.Field

    Set recordSet = New ADODB.recordSet
    recordSet.Open SQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
 
    If recordSet.State Then
        
        RefSheet.Cells(2, col).CopyFromRecordset recordSet
        
        For Each Field In recordSet.Fields
            Cells(1, col) = Field.Name
            col = col + 1
        Next Field
       
        Set recordSet = Nothing
    End If
End Function

Public Sub Get_Platform()
 
    Dim SQL_Platform As String, col As Long
    Dim Connected As Boolean
    Dim rg As Range
    
    Set rg = RefSheet.Range("A1").CurrentRegion
    rg.ClearContents
 
    'GETTING PLATFORM
    SQL_Platform = "SELECT DISTINCT Platform FROM ExposurePOR.dbo.POR ORDER BY Platform"
    col = 1
    
    Connected = ConnectToDB(dbAddress, uName, pWord)
    
    If Connected Then
        Query SQL_Platform, col
        Conn.Close
    Else
        MsgBox "Please connect to HP Remote Access and try opening again"
        'Exit Sub
        ActiveWorkbook.Close
    End If
       
End Sub

