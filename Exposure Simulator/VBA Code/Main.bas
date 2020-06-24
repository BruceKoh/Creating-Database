Attribute VB_Name = "Main"
Option Explicit

Sub user_main()
    
    Get_Platform
    
    DisplayUserList
    
End Sub

Sub DisplayUserList()
Attribute DisplayUserList.VB_ProcData.VB_Invoke_Func = "u\n14"

    Dim form As New Platform, rgsku As Range
    With form
        .Height = 500
        .Width = 930
        .Show
    End With
    
    If form.Cancelled = True Then
        MsgBox "The UserForm was cancelled."
        Exit Sub
    End If
   
    
End Sub

Sub user_done()
    
    Dim dict As New Dictionary
    
    Dim sh As Worksheet
    Set sh = ThisWorkbook.Worksheets("Quantity")
    
    
    Dim rg As Range, i As Long, rgout As Range, row As Long
    Dim sku As String, j As Long
    Set rg = Quan.Range("A1").CurrentRegion
    Set rgout = output.Range("A1").CurrentRegion
    
    Dim oUser As clsUser
      
    output.Range("H1") = "Quantity"
    If MsgBox("Is the quantity entered correct? ", vbYesNo) = vbYes Then
        For i = 2 To rg.Rows.Count
            Set oUser = New clsUser
            oUser.sku = rg.Cells(i, 1)
            oUser.quantity = rg.Cells(i, 2)
            oUser.mpa = rg.Cells(i, 3)
            dict.Add oUser.sku, oUser
        Next i
        'Remove duplicates

        Set rgout = output.Range("A1").CurrentRegion
        Dim key As Variant
        
        
        For Each key In dict.Keys
            For j = 2 To rgout.Rows.Count
                Set oUser = dict(key)
                sku = output.Cells(j, 2).Value
                If sku = key Then
                    output.Cells(j, 10) = oUser.quantity
'                    output.Cells(j, 10) = oUser.mpa
                End If
            Next j
        Next key

        output.Range("J1") = "User Quantity"
'       get_price
        multiply
        output.Range("P2") = Application.WorksheetFunction.Sum(output.Range("K:K"))
        output.Range("K1") = "Total Cost"
        output.Activate
    Else
        MsgBox "Please Enter Quantity Again!"
    End If

    Application.CutCopyMode = False
    On Error Resume Next
    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 1), Cells(j - 1, 11)), , xlYes).Name = _
        "Table1"
    'Range("Table1[#All]").Select
    ThisWorkbook.RefreshAll
    Worksheets("Output Summary").Select
    Cells(4, 1).Select
    
End Sub


Sub multiply()
    
    Dim multi As Variant
    Dim rg As Range, i As Long
    Set rg = output.Range("A1").CurrentRegion
    Set rg = rg.Resize(rg.Rows.Count - 1).Offset(1)
    multi = rg.Value2
    ReDim Preserve multi(1 To rg.Rows.Count, 1 To rg.Columns.Count + 1)
    
    For i = LBound(multi) To UBound(multi)
        multi(i, 11) = multi(i, 7) * multi(i, 8) * multi(i, 10)
    Next i
    
    output.Range("A2:K" & rg.Rows.Count + 1).Value = multi
    

End Sub
