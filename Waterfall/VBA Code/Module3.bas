Attribute VB_Name = "Module3"
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 80)
        .Transparency = 0
        .Solid
    End With
    With Selection.Format.Line
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 176, 80)
        .Transparency = 0
    End With
End Sub
