VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Platform 
   Caption         =   "Waterall UI"
   ClientHeight    =   7536
   ClientLeft      =   -120
   ClientTop       =   -495
   ClientWidth     =   10650
   OleObjectBlob   =   "Platform.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Platform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private m_POR As Collection
Private m_YEAR As Collection
Private m_base As String
Private m_compare As String
Private m_Cancelled As Boolean
Public Property Get Cancelled() As Boolean
    Cancelled = m_Cancelled
End Property
Property Get POR() As Collection
    Set POR = m_POR
End Property
Property Get year() As Collection
    Set year = m_YEAR
End Property
Private Function GetSelections() As Collection

    Dim collPOR As New Collection
    Dim i As Long
    
    'Go through each item in listbox
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) Then
            collPOR.Add ListBox1.List(i)
        End If
    Next i
    
    
    
    Set GetSelections = collPOR
    
End Function
Private Function GetSelectionsYEAR() As Collection

    Dim collYEAR As New Collection
    Dim i As Long
    
    'Go through each item in listbox
    For i = 0 To ListBox2.ListCount - 1
        If ListBox2.Selected(i) Then
            collYEAR.Add ListBox2.List(i)
        End If
    Next i
        
    Set GetSelectionsYEAR = collYEAR
    
End Function

Private Sub ButtonCancel_Click()
    ' Hide the Userform and set cancelled to true
    Hide
    m_Cancelled = True
End Sub

Private Sub ButtonOK_Click()

    Set m_POR = GetSelections
    
    Set m_YEAR = GetSelectionsYEAR

    Hide
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

    Dim rgplat As Range, rgplan As Range
    Set rgplat = RefSheet.Range("A1").CurrentRegion
    Set rgplan = RefSheet.Range("C1").CurrentRegion

    ListBox1.List = rgplat.Range("A2" & ":" & "A" & rgplat.Rows(rgplat.Rows.Count).Row).Value2
    ListBox2.List = RefSheet.Range("C2:C" & rgplan.Rows.Count).Value2


End Sub
