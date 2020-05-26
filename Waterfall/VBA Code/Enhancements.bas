Attribute VB_Name = "Enhancements"
Public Sub TurnOffFunctionality()
    Application.Calculation = xlCalculationManual
    Application.DisplayStatusBar = False
    Application.EnableEvents = False
    Application.ScreenUpdating = False
End Sub

Public Sub TurnOnFunctionality()
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayStatusBar = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Sub RefreshWaterfall()

Dim lastRow As Long
Dim WFsheet As Worksheet
Dim DashBsheet As Worksheet
Dim rowCount As Integer
Set WFsheet = Worksheets("PivotTable4")
Set DashBsheet = Worksheets("Waterfall")

WFsheet.Select
lastRow = Cells(Rows.Count, 2).End(xlUp).Row

If Cells(lastRow, 2).Value = "(blank)" Then
    lastRow = lastRow - 1
End If


'    With WFsheet.ListObjects("Table1")
'            If Not .DataBodyRange Is Nothing Then
'                .DataBodyRange.Delete
'            End If
'    End With
    
    With WFsheet
'            .Range(Cells(5, 6), Cells(lastRow, 6)).FormulaR1C1 = "=RC[-4]"
'            .Range(Cells(5, 7), Cells(lastRow, 7)).FormulaR1C1 = "=IF(ISBLANK(RC[-3]),RC[-4],IF(ISBLANK(R[1]C[-5]),RC[-4],RC[-3]))"
            rowCount = .Range(Cells(5, 2), Cells(lastRow, 2)).Count
    End With
    
 DashBsheet.Select
  
    With DashBsheet.ListObjects("Table3")
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With

    With DashBsheet
        .Range(Cells(5, 2), Cells(lastRow, 2)).FormulaR1C1 = "=PivotTable4!RC"
        .Range(Cells(5, 3), Cells(lastRow, 3)).FormulaR1C1 = "=IF(ISBLANK(PivotTable4!RC[1]),PivotTable4!RC,IF(ISBLANK(PivotTable4!R[1]C),PivotTable4!RC,PivotTable4!RC[1]))"
        .ChartObjects("Chart 4").Activate
            With ActiveChart
                .FullSeriesCollection(1).Points(1).IsTotal = True
                .FullSeriesCollection(1).Points(rowCount).IsTotal = True
                .ChartTitle.Caption = "Waterfall Chart by Platform"
            End With
    End With

    For iRow = 5 To lastRow
            
        PORVer = DashBsheet.Cells(iRow, 2).Value
        
        If iRow <> 5 And iRow <> lastRow Then
            With Worksheets("Pvt_PORvPOR+Ship")
                .Select
                .Range("C14").Select
                .PivotTables("PivotTable3").PivotFields("Planning_Wk").PivotItems(PORVer).Visible = False
            End With

        ElseIf iRow = 5 Then
        
            With Worksheets("Pvt_PORvPOR+Ship")
                .Select
                .Range("C14").Select
                .PivotTables("PivotTable3").PivotFields("Planning_Wk").PivotItems(PORVer).Visible = True
            End With
        
        ElseIf iRow = lastRow Then
            With Worksheets("Pvt_PORvPOR+Ship")
                .Select
                .Range("C14").Select
                .PivotTables("PivotTable3").PivotFields("Planning_Wk").PivotItems(PORVer).Visible = True
            End With
            With DashBsheet
                .Select
                .ChartObjects("Chart 6").Activate
                ActiveChart.FullSeriesCollection(1).ChartType = xlLine
            End With
        End If
        
    Next iRow
        
End Sub

Sub Change2Line()
Dim DashBsheet As Worksheet
Set DashBsheet = Worksheets("Waterfall")

    lastRow1 = Worksheets("Pvt_PORvPOR+Ship").Cells(Rows.Count, 3).End(xlUp).Row
    ChartSeriesCount = lastRow1 - 3
    
    If ChartSeriesCount = 2 Then
        With DashBsheet
            .Select
            .ChartObjects("Chart 6").Activate
            With ActiveChart
                seriesName1 = Trim(Right(.FullSeriesCollection(1).Name, 4))
                seriesName2 = Trim(Right(.FullSeriesCollection(2).Name, 4))
                seriesPrefix1 = Trim(Left(.FullSeriesCollection(1).Name, 7))
                seriesPrefix2 = Trim(Left(.FullSeriesCollection(2).Name, 7))
            
'    With Selection.Format.Line
'        .Visible = msoTrue
'        .ForeColor.RGB = RGB(0, 176, 80)
'        .Transparency = 0
'    End With
            
                If seriesPrefix1 <> seriesPrefix2 And seriesName1 = "POR" And seriesName2 = "SHIP" Then
                    With ActiveChart
                    'Stop
                        .FullSeriesCollection(1).ChartType = xlLine
                        .FullSeriesCollection(2).ChartType = xlColumnClustered
                        .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)  'green
                        .FullSeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 176, 80)  'green
                        .FullSeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 153, 0)  'orange
                    End With
                ElseIf seriesPrefix1 <> seriesPrefix2 And seriesName1 = "POR" And seriesName2 = "POR" Then
                    With ActiveChart
                        .FullSeriesCollection(1).ChartType = xlLine
                        .FullSeriesCollection(2).ChartType = xlColumnClustered
                        .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)  'green
                        .FullSeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 176, 80)  'green
                        .FullSeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 153, 102) 'bright orange
                    End With
                ElseIf seriesPrefix1 = seriesPrefix2 And seriesName1 = "POR" And seriesName2 = "SHIP" Then
                    With ActiveChart
                        .FullSeriesCollection(1).ChartType = xlColumnClustered
                        .FullSeriesCollection(2).ChartType = xlColumnClustered
                        .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 112, 192)  'blue
                        .FullSeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 153, 102) 'bright orange
                    End With
                ElseIf seriesName1 = "Ship" And seriesName2 = "SHIP" Then
                    With ActiveChart
                        .FullSeriesCollection(1).ChartType = xlColumnClustered
                        .FullSeriesCollection(2).ChartType = xlColumnClustered
                        .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 153, 102) 'bright orange
                        .FullSeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(255, 153, 102) 'bright orange
                    End With
                End If
                
                
                GoTo EndChangeToLine
            End With
        End With
        
    ElseIf ChartSeriesCount = 1 Then
        With DashBsheet
            .Select
            .ChartObjects("Chart 6").Activate
            With ActiveChart
                seriesName = Trim(Right(.FullSeriesCollection(1).Name, 4))
                If seriesName = "POR" Then
                    .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)  'green
                    .FullSeriesCollection(1).Format.Line.ForeColor.RGB = RGB(0, 176, 80)  'green
                ElseIf seriesName = "SHIP" Then
                    .FullSeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(255, 153, 102) 'bright orange
                End If
            End With
        End With
        
        GoTo EndChangeToLine
    End If
    
    For ChartSeriesCounter = 1 To ChartSeriesCount
    With DashBsheet
        .ChartObjects("Chart 6").Activate
        With ActiveChart
            seriesName = Trim(Right(.FullSeriesCollection(ChartSeriesCounter).Name, 4))
            seriesBase = Trim(Left(.FullSeriesCollection(1).Name, 7))
            seriesPrefix = Trim(Left(.FullSeriesCollection(ChartSeriesCounter).Name, 7))
        End With
    End With
    
        'Series 1
        If ChartSeriesCounter = 1 Then
            With DashBsheet
                .Select
                .ChartObjects("Chart 6").Activate
                With ActiveChart
                    .FullSeriesCollection(ChartSeriesCounter).ChartType = xlLine
                    
                    If Trim(Right(.FullSeriesCollection(ChartSeriesCounter).Name, 4)) = "POR" Then
                        .FullSeriesCollection(ChartSeriesCounter).Format.Fill.ForeColor.RGB = RGB(0, 176, 80)  'green
                        .FullSeriesCollection(ChartSeriesCounter).Format.Line.ForeColor.RGB = RGB(0, 176, 80)  'green
                    Else
                        .FullSeriesCollection(ChartSeriesCounter).Format.Fill.ForeColor.RGB = RGB(255, 153, 0)  'orange
                    End If
                End With
            End With
        Else
            With DashBsheet
                .Select
                .ChartObjects("Chart 6").Activate
                With ActiveChart
                    .FullSeriesCollection(ChartSeriesCounter).ChartType = xlColumnClustered
                    If Trim(Right(.FullSeriesCollection(ChartSeriesCounter).Name, 4)) = "POR" Then
                        .FullSeriesCollection(ChartSeriesCounter).Format.Fill.ForeColor.RGB = RGB(0, 112, 255) 'bright blue
                        .FullSeriesCollection(ChartSeriesCounter).Format.Line.Visible = msoFalse
                    Else
                        .FullSeriesCollection(ChartSeriesCounter).Format.Fill.ForeColor.RGB = RGB(255, 153, 102) 'bright orange
                        If seriesBase = seriesPrefix Then
                            With ActiveChart
                            .FullSeriesCollection(ChartSeriesCounter).ChartType = xlLine
                            End With
                        End If
                    End If
                End With
            End With
        
        End If
        
    Next ChartSeriesCounter

EndChangeToLine:
'''            With DashBsheet
'''                .Select
'''                .ChartObjects("Chart 6").Activate
'''                With ActiveChart
'''                    .FullSeriesCollection(1).ChartType = xlLine
'''
'''                    .FullSeriesCollection(2).Select
'''                    With Selection.Format.Fill
'''                        .Visible = msoTrue
'''                        .ForeColor.RGB = RGB(0, 112, 192)
'''                        .Transparency = 0
'''                        .Solid
'''                    End With
'''
'''                    .FullSeriesCollection(3).Select
'''                    With Selection.Format.Fill
'''                        .Visible = msoTrue
'''                        .ForeColor.ObjectThemeColor = msoThemeColorAccent2
'''                        .ForeColor.TintAndShade = 0
'''                        .ForeColor.Brightness = 0
'''                        .Transparency = 0
'''                        .Solid
'''                    End With
'''                End With
'''            End With
Cells(1, 1).Select


End Sub
