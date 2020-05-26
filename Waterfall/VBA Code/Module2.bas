Attribute VB_Name = "Module2"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    ActiveSheet.ChartObjects("Chart 6").Activate
    ActiveChart.FullSeriesCollection(2).Select
    ActiveWindow.ScrollColumn = 2
    ActiveChart.FullSeriesCollection(2).Points(43).Select
    Range("O32").Select
    ActiveSheet.ChartObjects("Chart 6").Activate
    ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 112, 192)
        .Transparency = 0
        .Solid
    End With
    Range("Q33").Select
End Sub
