Attribute VB_Name = "SQL_GET_PLATFORM_PLAN"
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

Public Sub Run()
 
    Dim SQL As String, SQL2 As String, col As Long, col2 As Long
    Dim Connected As Boolean
    Dim rg1 As Range, rg2 As Range
    
    Set rg1 = RefSheet.Range("A1").CurrentRegion
    Set rg2 = RefSheet.Range("C1").CurrentRegion
    rg1.ClearContents
    rg2.ClearContents
 
    'GETTING PLATFORM
    SQL = "SELECT DISTINCT SHIP.Platform AS PLATFORM FROM [SHIPMENT].dbo.SHIPMENT AS SHIP ORDER BY Platform"
    col = 1
    'GETTING PLANNING_WK
    SQL2 = "SELECT DISTINCT Planning_Wk FROM FULLSHIPVPOR.dbo.FULLSHIPVPOR ORDER BY Planning_Wk"
    Connected = ConnectToDB(dbAddress, uName, pWord)
    col2 = 3
    
    If Connected Then
        Query SQL, col
        Query SQL2, col2
        Conn.Close
    Else
        MsgBox "Please ensure that you are connected to HP Remote Access"
    End If
    
    cleantime
    
End Sub

Private Sub clean()

    Dim c As Range, rngConstants As Range

    On Error Resume Next

    Set rngConstants = RefSheet.Range("A1").CurrentRegion

    On Error GoTo 0

    If Not rngConstants Is Nothing Then

        'optimize performance

        TurnOffFunctionality

        'trim cells incl char 160

        For Each c In rngConstants

            c.Value = Trim$(Application.clean(Replace(c.Value, Chr(160), "")))

        Next c

        'reset settings

        TurnOnFunctionality
        
    End If
    
End Sub

Private Sub cleantime()

'PURPOSE: Determine how many seconds it took for code to completely run
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault

Dim StartTime As Double
Dim SecondsElapsed As Double

'Remember time when macro starts
StartTime = Timer
  
clean

End Sub
