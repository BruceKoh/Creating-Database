Attribute VB_Name = "DYNAMICSQL"
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

'DYNAMIC FROM HERE ONWARDS

Public Sub SQLPORBASE(ByRef Platform As Collection, ByRef year As Collection)
    
    Dim SQL As String
    Dim Connected As Boolean, col As Long, coltitle As Long
    Dim rerg As Range, rg As Range
    Set rerg = RefSheet.Range("E1").CurrentRegion
    rerg.ClearContents
    
    Dim P As Variant
    For Each P In Platform
        coltitle = 5
        col = 5
        'GETTING BASE POR DATA
        SQL = "SELECT Planning_Wk,YYYYWW,Qty,Platform,MPA,Region,QtyType FROM FULLSHIPVPOR.dbo.FULLSHIPVPOR WHERE Platform = " & " '" & P & _
        "'" & " AND QtyType = 'POR' AND Planning_Wk >= " & "'" & year(1) & "'" & " AND Planning_Wk <= " & "'" & year(year.Count) & "'"
 
        Connected = ConnectToDB(dbAddress, uName, pWord)
        Set rg = RefSheet.Range("E1").CurrentRegion
        
        If Connected Then
            QueryCol SQL, col, coltitle, rg
            Conn.Close
        Else
            MsgBox "We have a problem!"
        End If
    Next P
    
End Sub

Public Sub SQLSHIPMENT(ByRef Platform As Collection, ByRef year As Collection)
    
    Dim SQL As String, YYYYFIRST As Long, YYYYLAST As Long
    Dim Connected As Boolean, col As Long, coltitle As Long
    Dim rerg As Range, rg As Range
    Set rerg = RefSheet.Range("M1").CurrentRegion
    rerg.ClearContents

    YYYYFIRST = Replace(year(1), "W", "")
    YYYYLAST = Replace(year(year.Count), "W", "")
    Dim P As Variant
    For Each P In Platform
        coltitle = 13
        col = 13
        'GETTING SHIPMENT DATA
        SQL = "SELECT Planning_Wk,YYYYWW,Qty,Platform,MPA,Region,QtyType FROM FULLSHIPVPOR.dbo.FULLSHIPVPOR WHERE platform = " & " '" & P & _
        "'" & " AND QtyType = 'SHIP' " & " AND Planning_Wk >= " & " '" & year(1) & "'" & " AND Planning_Wk <= " & " '" & year(year.Count) & "'" & _
        " AND YYYYWW > " & YYYYFIRST & " AND YYYYWW < " & YYYYLAST
        

        Connected = ConnectToDB(dbAddress, uName, pWord)
        Set rg = RefSheet.Range("M1").CurrentRegion

        If Connected Then
            QueryCol SQL, col, coltitle, rg
            Conn.Close
        Else
            MsgBox "We have a problem!"
        End If
    Next P
    
End Sub

Private Function QueryCol(SQL As String, col As Long, coltitle As Long, rg As Range)
 
    Dim recordSet As ADODB.recordSet
    Dim Field As ADODB.Field
    
    Set recordSet = New ADODB.recordSet
    recordSet.Open SQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
 
    If recordSet.State Then
        
        For Each Field In recordSet.Fields
            RefSheet.Cells(1, coltitle) = Field.Name
            coltitle = coltitle + 1
        Next Field
        
        RefSheet.Cells(rg.Rows.Count + 1, col).CopyFromRecordset recordSet
        
        Set recordSet = Nothing
    End If
    
End Function



