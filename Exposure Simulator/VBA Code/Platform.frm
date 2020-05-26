VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Platform 
   Caption         =   "Exposure Simulator UI"
   ClientHeight    =   10170
   ClientLeft      =   -120
   ClientTop       =   -495
   ClientWidth     =   22545
   OleObjectBlob   =   "Platform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Platform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_plat_base As Collection
Private m_region_base As Collection
Private m_sku_base As Collection
Private m_Cancelled As Boolean
Private m_sku As Boolean
Private m_plat_compare As Collection
Private m_region_compare As Collection
Public Property Get sku() As Boolean
    sku = m_sku
End Property
Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancelled
End Property
Property Get plat_base() As Collection
    Set plat_base = m_plat_base
End Property
Property Get Region_B() As Collection
    Set Region_B = m_region_base
End Property
Property Get Sku_B() As Collection
    Set Sku_B = m_sku_base
End Property
Property Get plat_compare() As Collection
    Set plat_compare = m_plat_compare
End Property
Property Get Region_C() As Collection
    Set Region_C = m_region_compare
End Property
Private Function GetSelectionsPlatBase() As Collection

    Dim collPlatBase As New Collection
    Dim i As Long
    
    'Go through each item in listbox
    For i = 0 To Platform_Base.ListCount - 1
        If Platform_Base.Selected(i) Then
            collPlatBase.Add Platform_Base.List(i)
        End If
    Next i
    
    Set GetSelectionsPlatBase = collPlatBase
    
End Function
Private Function GetSelectionsREGIONBase() As Collection

    Dim collREGIONBase As New Collection
    Dim i As Long
    
    'Go through each item in listbox
    For i = 0 To Region_Base.ListCount - 1
        If Region_Base.Selected(i) Then
            collREGIONBase.Add Region_Base.List(i)
        End If
    Next i
    
    Set GetSelectionsREGIONBase = collREGIONBase
    
End Function
Private Function GetSelectionsREGIONCompare() As Collection

    Dim collREGIONCompare As New Collection
    Dim i As Long
    
    'Go through each item in listbox
    For i = 0 To Region_Compare.ListCount - 1
        If Region_Compare.Selected(i) Then
            collREGIONCompare.Add Region_Compare.List(i)
        End If
    Next i
    
    Set GetSelectionsREGIONCompare = collREGIONCompare
    
End Function
Private Function GetSelectionsSKUBase() As Collection

    Dim collskubase As New Collection
    Dim i As Long
    
    'Go through each item in listbox
    For i = 0 To Base_SKU.ListCount - 1
        If Base_SKU.Selected(i) Then
            collskubase.Add Base_SKU.List(i, 0)
        End If
    Next i
    
    Set GetSelectionsSKUBase = collskubase
    
End Function
Private Function GetSelectionsPlatCompare() As Collection

    Dim collPlatCompare As New Collection
    Dim i As Long
    
    'Go through each item in listbox
    For i = 0 To Platform_Compare.ListCount - 1
        If Platform_Compare.Selected(i) Then
            collPlatCompare.Add Platform_Compare.List(i)
        End If
    Next i
    
    Set GetSelectionsPlatCompare = collPlatCompare
    
End Function

Private Sub Button_Quantity_Click()

    Dim collskucompare As New Collection
    Dim collskubase As New Collection
    Dim collmpa As New Collection
    
    Dim i As Long, j As Long, k As Long
          
    'Base SKU
    For j = 0 To Base_SKU.ListCount - 1
        If Base_SKU.Selected(j) = True Then
            collskubase.Add Base_SKU.List(j, 0)
        End If
    Next j
    
    For k = 0 To Base_SKU.ListCount - 1
        If Base_SKU.Selected(k) = True Then
            collmpa.Add Base_SKU.List(k, 1)
        End If
    Next k
    
    'Compared sku
    For i = 0 To Compare_SKU.ListCount - 1
        If Compare_SKU.Selected(i) Then
            collskucompare.Add Compare_SKU.List(i)
        End If
    Next i

    
    Dim item As Variant, mpa As Variant, row As Long, rg As Range, rowmpa As Long
    Set rg = Quan.Range("A1").CurrentRegion
    rg.Cells.ClearContents
    
'    SQLSKU_MPA collskubase
    
    Hide
    row = 2
    For Each item In collskubase
        Quan.Cells(row, 1).Value = item
        row = row + 1
    Next
    
    rowmpa = 2
    For Each mpa In collmpa
        Quan.Cells(rowmpa, 3).Value = mpa
        rowmpa = rowmpa + 1
    Next

    Quan.Activate
    Quan.Range("A1") = "SKU Selected"
    Quan.Range("B1") = "Enter Quantity"
    Quan.Range("C1") = "MPA"
    
    SQLSKUCOMPARE collskubase, collskucompare
    
    
    
End Sub

Private Sub ButtonCancel_Click()
    ' Hide the Userform and set cancelled to true
    Hide
    m_Cancelled = True
End Sub

Private Sub ButtonRegion_Click()

    Set m_plat_base = GetSelectionsPlatBase

    SQLPLATFORM_REGION plat_base
    
    Dim region_range As Range
    Set region_range = RefSheet.Range("C1").CurrentRegion
    If region_range.Rows.Count = 2 Then
        Region_Base.Clear
        Region_Base.AddItem RefSheet.Range("C2").Value2
    ElseIf IsEmpty(RefSheet.Range("C2")) Then
        MsgBox ("Invalid please select again")
    Else
        Set region_range = RefSheet.Range("C1:C" & region_range.Rows.Count)
        region_range.RemoveDuplicates Columns:=Array(1), Header:=xlYes
        Region_Base.List = RefSheet.Range("C2:C" & region_range.Rows.Count).Value2
    End If
    
End Sub

Private Sub ButtonRegion_Compare_Click()

    Set m_plat_compare = GetSelectionsPlatCompare

    SQLPLATFORM_REGION_COMPARE plat_compare
    
    Dim region_range As Range
    Set region_range = RefSheet.Range("H1").CurrentRegion
    If region_range.Rows.Count = 2 Then
        Region_Compare.Clear
        Region_Compare.AddItem RefSheet.Range("H2").Value2
    ElseIf IsEmpty(RefSheet.Range("H2")) Then
        MsgBox ("Invalid please select again")
    Else
        Set region_range = RefSheet.Range("H1:H" & region_range.Rows.Count)
        region_range.RemoveDuplicates Columns:=Array(1), Header:=xlYes
        Region_Compare.List = RefSheet.Range("H2:H" & region_range.Rows.Count).Value2
    End If
    
End Sub

Private Sub ButtonSKU_Click()
   
    Set m_plat_base = GetSelectionsPlatBase
    Set m_region_base = GetSelectionsREGIONBase

    SQLPLATFORM_REGION_SKU plat_base, Region_B

    Dim sku_range As Range
    Set sku_range = RefSheet.Range("E1").CurrentRegion
'    If sku_range.Rows.Count = 2 Then
'        Base_SKU.Clear
'        Base_SKU.AddItem.List(0, 0) = RefSheet.Range("E2").Value2
'    ElseIf IsEmpty(RefSheet.Range("E2")) Then
'        MsgBox ("Invalid please select again")
'    Else
'        Set sku_range = RefSheet.Range("E1:E" & sku_range.Rows.Count)
'        sku_range.RemoveDuplicates Columns:=Array(1), Header:=xlYes
'        Base_SKU.List(1, 0) = RefSheet.Range("E2:E" & sku_range.Rows.Count).Value2
'    End If
    Dim i As Long, mpa As String, listrow As Long
    listrow = 0
    With Base_SKU
        .Clear
        For i = 0 To sku_range.Rows.Count
            mpa = RefSheet.Range("F" & i + 2).Value2
            If CheckBox5.Value = True And mpa = "Foxconn ChongQing" Then
                .AddItem
                .List(listrow, 0) = RefSheet.Range("E" & i + 2).Value2
                .List(listrow, 1) = RefSheet.Range("F" & i + 2).Value2
                listrow = listrow + 1
            ElseIf CheckBox6.Value = True And mpa = "Flex PTP Malasya" Then
                .AddItem
                .List(listrow, 0) = RefSheet.Range("E" & i + 2).Value2
                .List(listrow, 1) = RefSheet.Range("F" & i + 2).Value2
                listrow = listrow + 1
            ElseIf CheckBox7.Value = True And mpa = "Flex Zhuhai" Then
                .AddItem
                .List(listrow, 0) = RefSheet.Range("E" & i + 2).Value2
                .List(listrow, 1) = RefSheet.Range("F" & i + 2).Value2
                listrow = listrow + 1
            ElseIf CheckBox8.Value = True And mpa = "NKG Yue Yang" Then
                .AddItem
                .List(listrow, 0) = RefSheet.Range("E" & i + 2).Value2
                .List(listrow, 1) = RefSheet.Range("F" & i + 2).Value2
                listrow = listrow + 1
            ElseIf CheckBox9.Value = True And mpa = "NKG Thailand" Then
                .AddItem
                .List(listrow, 0) = RefSheet.Range("E" & i + 2).Value2
                .List(listrow, 1) = RefSheet.Range("F" & i + 2).Value2
                listrow = listrow + 1
            End If
            
        Next i
    End With
    
End Sub

Private Sub ButtonSKU_Compare_Click()
    
    Set m_plat_compare = GetSelectionsPlatCompare
    Set m_region_compare = GetSelectionsREGIONCompare

    SQLPLATFORM_REGION_SKU_COMPARE plat_compare, Region_C
    
    Dim sku_range As Range
    Set sku_range = RefSheet.Range("J1").CurrentRegion
    If sku_range.Rows.Count = 2 Then
        Compare_SKU.Clear
        Compare_SKU.AddItem RefSheet.Range("J2").Value2
    ElseIf IsEmpty(RefSheet.Range("J2")) Then
        MsgBox ("Invalid please select again")
    Else
        Set sku_range = RefSheet.Range("J1:J" & sku_range.Rows.Count)
        sku_range.RemoveDuplicates Columns:=Array(1), Header:=xlYes
        Compare_SKU.List = RefSheet.Range("J2:J" & sku_range.Rows.Count).Value2
    End If
    
End Sub

Private Sub CheckBox1_Click()

    Dim i As Long
    If CheckBox1.Value = True Then
        For i = 0 To Region_Base.ListCount - 1
            Region_Base.Selected(i) = True
        Next i
    Else
        For i = 0 To Region_Base.ListCount - 1
            Region_Base.Selected(i) = False
        Next i
    End If

End Sub

Private Sub CheckBox2_Click()
    
    Dim i As Long
    If CheckBox2.Value = True Then
        For i = 0 To Base_SKU.ListCount - 1
            Base_SKU.Selected(i) = True
        Next i
    Else
        For i = 0 To Base_SKU.ListCount - 1
            Base_SKU.Selected(i) = False
        Next i
    End If
    
End Sub

Private Sub CheckBox3_Click()
    
    Dim i As Long
    If CheckBox3.Value = True Then
        For i = 0 To Region_Compare.ListCount - 1
            Region_Compare.Selected(i) = True
        Next i
    Else
        For i = 0 To Region_Compare.ListCount - 1
            Region_Compare.Selected(i) = False
        Next i
    End If
    
End Sub

Private Sub CheckBox4_Click()
    
    Dim i As Long
    If CheckBox4.Value = True Then
        For i = 0 To Compare_SKU.ListCount - 1
            Compare_SKU.Selected(i) = True
        Next i
    Else
        For i = 0 To Compare_SKU.ListCount - 1
            Compare_SKU.Selected(i) = False
        Next i
    End If
    
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer _
                                       , CloseMode As Integer)
    
    ' Prevent the form being unloaded
    If CloseMode = vbFormControlMenu Then Cancel = True
    
    ' Hide the Userform and set cancelled to true
    Hide
    m_Cancelled = True
    
End Sub

Private Sub UserForm_Initialize()

    Dim rgplat As Range, rgplatcompare As Range
    Set rgplat = RefSheet.Range("A1").CurrentRegion
    Platform_Base.List = rgplat.Range("A2" & ":" & "A" & rgplat.Rows(rgplat.Rows.Count).row).Value2
    Platform_Compare.List = rgplat.Range("A2" & ":" & "A" & rgplat.Rows(rgplat.Rows.Count).row).Value2
    Base_SKU.ColumnCount = 2
    
 
End Sub
