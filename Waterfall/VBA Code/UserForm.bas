Attribute VB_Name = "UserForm"
Option Explicit

Sub DisplayUserList()
Attribute DisplayUserList.VB_ProcData.VB_Invoke_Func = "u\n14"

    Dim form As New Platform
    
    'Run
    
    form.Height = 400
    form.Width = 530
    
    form.Show
    
    
    TurnOffFunctionality
    
    If form.Cancelled = True Then
        MsgBox "The UserForm was cancelled."
        Exit Sub
    End If
   
    SQLPORBASE form.POR, form.year

    SQLSHIPMENT form.POR, form.year

    MakeTable
    Worksheets("RefSheet").Select
  With Range("U1")
    .Parent.ListObjects.Add(xlSrcRange, Range(.End(xlDown), .End(xlToRight)), , xlYes).Name = "Table1"
  End With
  
  Dim lastRow As Long
  lastRow = Cells(Rows.Count, 21).End(xlUp).Row
  Range(Cells(2, 28), Cells(lastRow, 28)).FormulaR1C1 = "=Left(RC[-6],4)"
  Cells(1, 28).Value = "Year"
  
  
  
'    Dim lastRow As Long
'    Dim PORShp_rg As Range
'    Dim RefSht As Worksheet
'
'    Set RefSht = Worksheets("RefSheet")
'    RefSht.Select
'    lastRow = Cells(Rows.Count, 21).End(xlUp).Row
'    Set PORShp_rg = Range(Cells(1, 21), Cells(27, lastRow))
'    Application.CutCopyMode = False
'    Cells(1, 21).Select
'    Range(Cells(1, 21), Cells(27, lastRow)).Select
'    ActiveSheet.ListObjects.Add(xlSrcRange, Range(Cells(1, 21), Cells(27, lastRow)), , xlYes).Name _
'        = "Table1"
'    Range("Table1[#All]").Select
    
    pivotworksheet
    
    RefreshWaterfall
    
    TurnOnFunctionality

    Unload form
    Set form = Nothing
    
'    PrintCollection form.POR
    
End Sub

Public Sub PrintCollection(ByRef coll As Collection)
    
    Debug.Print "The user selected the following POR:"
    Dim v As Variant
    For Each v In coll
        Debug.Print v
    Next
    
End Sub
