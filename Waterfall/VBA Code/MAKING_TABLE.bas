Attribute VB_Name = "MAKING_TABLE"
Option Explicit

Sub MakeTable()
    
    Dim rgpor As Range, rgshipment As Range, rg_clear As Range
    
    Set rg_clear = RefSheet.Range("U1").CurrentRegion
    
    rg_clear.ClearContents
    
    Set rgpor = RefSheet.Range("E1").CurrentRegion
    
    rgpor.Copy
    
    RefSheet.Range("U1").PasteSpecial xlPasteValues
    
    Set rgshipment = RefSheet.Range("M1").CurrentRegion
    
    Set rgshipment = rgshipment.Resize(rgshipment.Rows.Count - 1).Offset(1)
    
    rgshipment.Copy
    
    RefSheet.Range("U" & rgpor.Rows.Count + 1).PasteSpecial xlPasteValues
    
End Sub

Sub pivotworksheet()

    Dim pt As PivotTable
    
    For Each pt In Pivot.PivotTables
    
        pt.RefreshTable
        
    Next pt
    
End Sub
