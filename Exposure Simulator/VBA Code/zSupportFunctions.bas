Attribute VB_Name = "zSupportFunctions"
Sub SortListBox(oLb As MSForms.ListBox, sCol As Integer, sType As Integer, sDir As Integer)
 
 Dim vaItems As Variant
 Dim i As Long, j As Long
 Dim c As Integer
 Dim vTemp As Variant
 
 'Put the items in a variant array
  vaItems = oLb.List
 
 'Sort the Array Alphabetically(1)
 If sType = 1 Then
 For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1
 For j = i + 1 To UBound(vaItems, 1)
 'Sort Ascending (1)
 If sDir = 1 Then
 If vaItems(i, sCol) > vaItems(j, sCol) Then
 For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
 On Error GoTo endMe
 vTemp = vaItems(i, c)
 vaItems(i, c) = vaItems(j, c)
 vaItems(j, c) = vTemp
 Next c
 End If
 'Sort Descending (2)
 ElseIf sDir = 2 Then
 If vaItems(i, sCol) < vaItems(j, sCol) Then
 For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
 vTemp = vaItems(i, c)
 vaItems(i, c) = vaItems(j, c)
 vaItems(j, c) = vTemp
 Next c
 End If
 End If
 
 Next j
 Next i
 'Sort the Array Numerically(2)
 '(Substitute CInt with another conversion type (CLng, CDec, etc.) depending on type of numbers in the column)
 ElseIf sType = 2 Then
 For i = LBound(vaItems, 1) To UBound(vaItems, 1) - 1
 For j = i + 1 To UBound(vaItems, 1)
 'Sort Ascending (1)
 If sDir = 1 Then
 If CInt(vaItems(i, sCol)) > CInt(vaItems(j, sCol)) Then
 For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
 vTemp = vaItems(i, c)
 vaItems(i, c) = vaItems(j, c)
 vaItems(j, c) = vTemp
 Next c
 End If
 'Sort Descending (2)
 ElseIf sDir = 2 Then
 If CInt(vaItems(i, sCol)) < CInt(vaItems(j, sCol)) Then
 For c = 0 To oLb.ColumnCount - 1 'Allows sorting of multi-column ListBoxes
 vTemp = vaItems(i, c)
 vaItems(i, c) = vaItems(j, c)
 vaItems(j, c) = vTemp
 Next c
 End If
 End If
 
 Next j
 Next i
 End If
 
endMe:
 
 'Set the list to the array
 oLb.List = vaItems
End Sub
