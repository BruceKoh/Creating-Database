Attribute VB_Name = "DynamicSQL"
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

Public Sub SQLPLATFORM_REGION(ByRef Platform As Collection)
    
    Dim SQL As String
    Dim Connected As Boolean, col As Long, coltitle As Long
    Dim rerg As Range, rg As Range
    Set rerg = RefSheet.Range("C1").CurrentRegion
    rerg.ClearContents
    
    Dim P As Variant
    For Each P In Platform
        coltitle = 3
        col = 3
        'GETTING BASE POR DATA
        SQL = "SELECT DISTINCT Region FROM ExposurePOR.dbo.POR WHERE Platform = " & " '" & P & "'"
 
        Connected = ConnectToDB(dbAddress, uName, pWord)
        Set rg = RefSheet.Range("C1").CurrentRegion
        
        If Connected Then
            QueryCol SQL, col, coltitle, rg
            Conn.Close
        Else
        MsgBox "Please connect to HP Remote Access"
        Exit Sub
        End If
    Next P
    
End Sub
Public Sub SQLPLATFORM_REGION_COMPARE(ByRef Platform As Collection)
    
    Dim SQL As String
    Dim Connected As Boolean, col As Long, coltitle As Long
    Dim rerg As Range, rg As Range
    Set rerg = RefSheet.Range("H1").CurrentRegion
    rerg.ClearContents
    
    Dim P As Variant
    For Each P In Platform
        coltitle = 8
        col = 8
        'GETTING COMPARE POR DATA
        SQL = "SELECT DISTINCT Region FROM ExposurePOR.dbo.POR WHERE Platform = " & " '" & P & "'"
 
        Connected = ConnectToDB(dbAddress, uName, pWord)
        Set rg = RefSheet.Range("H1").CurrentRegion
        
        If Connected Then
            QueryCol SQL, col, coltitle, rg
            Conn.Close
        Else
        MsgBox "Please connect to HP Remote Access"
        Exit Sub
        End If
    Next P
    
End Sub
Public Sub SQLMPA(ByRef Platform As Collection, ByRef region As Collection)
    
    Dim SQL As String
    Dim Connected As Boolean, col As Long, coltitle As Long
    Dim rerg As Range, rg As Range
    Set rerg = RefSheet.Range("L1").CurrentRegion
    rerg.ClearContents
    
    Dim P As Variant, R As Variant
    For Each P In Platform
        For Each R In region
            coltitle = 12
            col = 12
            'GETTING COMPARE POR DATA
            SQL = "SELECT DISTINCT MPA FROM ExposurePOR.dbo.POR WHERE Platform = " & " '" & P & "'" & _
            " AND Region = '" & R & "'"
 
            Connected = ConnectToDB(dbAddress, uName, pWord)
            Set rg = RefSheet.Range("L1").CurrentRegion
        
            If Connected Then
                QueryCol SQL, col, coltitle, rg
                Conn.Close
            Else
                MsgBox "Please connect to HP Remote Access"
                Exit Sub
            End If
        Next R
    Next P
    
End Sub

Public Sub SQLPLATFORM_REGION_SKU(ByRef Platform As Collection, ByRef region As Collection)
    
    Dim SQL As String
    Dim Connected As Boolean, col As Long, coltitle As Long
    Dim rerg As Range, rg As Range
    Set rerg = RefSheet.Range("E1").CurrentRegion
    rerg.ClearContents
    
    Dim P As Variant, R As Variant
    For Each P In Platform
        For Each R In region
            coltitle = 5
            col = 5
            'GETTING BASE POR DATA
            SQL = "SELECT DISTINCT SKU, MPA FROM ExposurePOR.dbo.POR WHERE Platform = " & " '" & P & "'" & _
            " AND REGION = '" & R & "'"
 
            Connected = ConnectToDB(dbAddress, uName, pWord)
            Set rg = RefSheet.Range("E1").CurrentRegion
        
            If Connected Then
                QueryCol SQL, col, coltitle, rg
                Conn.Close
            Else
                MsgBox "Please connect to HP Remote Access"
                Exit Sub
            End If
        Next R
    Next P
    
End Sub
Public Sub SQLPLATFORM_REGION_SKU_COMPARE(ByRef Platform As Collection, ByRef region As Collection)
    
    Dim SQL As String
    Dim Connected As Boolean, col As Long, coltitle As Long
    Dim rerg As Range, rg As Range
    Set rerg = RefSheet.Range("J1").CurrentRegion
    rerg.ClearContents
    
    Dim P As Variant, R As Variant
    For Each P In Platform
        For Each R In region
            coltitle = 10
            col = 10
            'GETTING BASE POR DATA
            SQL = "SELECT DISTINCT SKU FROM ExposurePOR.dbo.POR WHERE Platform = " & " '" & P & "'" & _
            " AND REGION = '" & R & "'"
 
            Connected = ConnectToDB(dbAddress, uName, pWord)
            Set rg = RefSheet.Range("J1").CurrentRegion
        
            If Connected Then
                QueryCol SQL, col, coltitle, rg
                Conn.Close
            Else
                MsgBox "Please connect to HP Remote Access"
                Exit Sub
            End If
        Next R
    Next P
    
End Sub

Sub SQLSKUCOMPARE(ByRef SKU_Base As Collection, ByRef SKU_Compare As Collection)

    Dim SQL As String
    Dim Connected As Boolean, col As Long, coltitle As Long
    Dim rerg As Range, rg As Range
    Set rerg = output.Range("A1").CurrentRegion
    rerg.ClearContents
    
    Dim mpa As String, row As Long, j As Long
    row = 2

    Dim Sku_B As Variant, Sku_C As Variant
    For Each Sku_B In SKU_Base
        mpa = Quan.Range("C" & row).Value2
        j = 1
        For Each Sku_C In SKU_Compare
            coltitle = 1
            col = 1
            'Compare Table
            If j = 1 Then
                SQL = "SELECT DISTINCT t1.Owner,t1.SKU,t1.PartRev,t1.Category,t1.Component,t1.Description,t1.[Per Rate], t1.Cost, t1.MPA FROM " & _
                "(SELECT * FROM ExposureSim.dbo.BOMParts WHERE SKU = '" & Sku_B & "'" & " AND MPA = '" & mpa & "'" & _
                ") AS t1 FULL OUTER JOIN (SELECT * FROM ExposureSim.dbo.BOMParts WHERE SKU = '" & Sku_C & "'" & _
                ") AS t2" & " ON (t1.Component = t2.Component) WHERE t2.Owner IS NULL"
            Else
                SQL = SQL & " UNION " & "SELECT DISTINCT t1.Owner,t1.SKU,t1.PartRev,t1.Category,t1.Component,t1.Description,t1.[Per Rate], t1.Cost, t1.MPA FROM " & _
                "(SELECT * FROM ExposureSim.dbo.BOMParts WHERE SKU = '" & Sku_B & "'" & " AND MPA = '" & mpa & "'" & _
                ") AS t1 FULL OUTER JOIN (SELECT * FROM ExposureSim.dbo.BOMParts WHERE SKU = '" & Sku_C & "'" & _
                ") AS t2" & " ON (t1.Component = t2.Component) WHERE t2.Owner IS NULL"
            End If
            j = j + 1
        Next Sku_C
        
        Connected = ConnectToDB(dbAddress, uName, pWord)
        Set rg = output.Range("A1").CurrentRegion

        If Connected Then
            QueryOut SQL, col, coltitle, rg
            Conn.Close
        Else
            MsgBox "Please connect to HP Remote Access"
            Exit Sub
        End If
        row = row + 1
    Next Sku_B

End Sub
Sub SQLSKU_MPA(ByRef SKU_Base As Collection)

    Dim SQL As String
    Dim Connected As Boolean, col As Long, coltitle As Long
    Dim rerg As Range, rg As Range
    Set rerg = Quan.Range("A1").CurrentRegion
    Set rerg = Quan.Range("C2:C" & rerg.Rows.Count)
    rerg.ClearContents

    Dim Sku_B As Variant
    For Each Sku_B In SKU_Base
        coltitle = 3
        col = 3
        'Compare Table
        SQL = "SELECT DISTINCT MPA FROM ExposurePOR.dbo.POR WHERE SKU = '" & Sku_B & "'"

        Connected = ConnectToDB(dbAddress, uName, pWord)
        Set rg = Quan.Range("C1").CurrentRegion

        If Connected Then
            QueryQuan SQL, col, coltitle, rg
            Conn.Close
        Else
            MsgBox "Please connect to HP Remote Access"
            Exit Sub
        End If

    Next Sku_B

End Sub

Sub get_price()

    Dim SQL As String
    Dim Connected As Boolean, col As Long, coltitle As Long
    Dim rerg As Range, rg As Range
    Set rg = output.Range("A1").CurrentRegion
    Set rerg = output.Range("K2:K" & rg.Rows.Count)
    rerg.ClearContents
    
    Dim comp As String, mpa As String, i As Long, row As Long
    row = 2
    For i = 2 To rg.Rows.Count
        comp = rg.Range("E" & i).Value2
        mpa = rg.Range("J" & i).Value2
        coltitle = 11
        col = 11
        SQL = "SELECT Price FROM Materials.dbo.MatMaster WHERE HPPN = '" & comp & _
        "' AND MPA = '" & mpa & "'"
        
        Connected = ConnectToDB(dbAddress, uName, pWord)
        If Connected Then
            Queryprice SQL, col, coltitle, row
            Conn.Close
        Else
            MsgBox "Please connect to HP Remote Access"
            Exit Sub
        End If
        
        row = row + 1
        
    Next i
    
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

Private Function QueryOut(SQL As String, col As Long, coltitle As Long, rg As Range)
 
    Dim recordSet As ADODB.recordSet
    Dim Field As ADODB.Field
    
    Set recordSet = New ADODB.recordSet
    recordSet.Open SQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
 
    If recordSet.State Then
        
        For Each Field In recordSet.Fields
            output.Cells(1, coltitle) = Field.Name
            coltitle = coltitle + 1
        Next Field
        
        output.Cells(rg.Rows.Count + 1, col).CopyFromRecordset recordSet
        
        Set recordSet = Nothing
    End If
    
End Function

Private Function QueryQuan(SQL As String, col As Long, coltitle As Long, rg As Range)
 
    Dim recordSet As ADODB.recordSet
    Dim Field As ADODB.Field
    
    Set recordSet = New ADODB.recordSet
    recordSet.Open SQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
 
    If recordSet.State Then
        
        For Each Field In recordSet.Fields
            Quan.Cells(1, coltitle) = Field.Name
            coltitle = coltitle + 1
        Next Field
        
        Quan.Cells(rg.Rows.Count + 1, col).CopyFromRecordset recordSet
        
        Set recordSet = Nothing
    End If
    
End Function

Private Function Queryprice(SQL As String, col As Long, coltitle As Long, row As Long)
 
    Dim recordSet As ADODB.recordSet
    Dim Field As ADODB.Field
    
    Set recordSet = New ADODB.recordSet
    recordSet.Open SQL, Conn, adOpenStatic, adLockReadOnly, adCmdText
 
    If recordSet.State Then
        
        For Each Field In recordSet.Fields
            output.Cells(1, coltitle) = Field.Name
            coltitle = coltitle + 1
        Next Field
        
        output.Cells(row, col).CopyFromRecordset recordSet
        
        Set recordSet = Nothing
    End If
    
End Function
