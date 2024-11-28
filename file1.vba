Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim headerRange As Range
    Dim cell As Range
    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set headerRange = ws.Range("A1").CurrentRegion.Rows(1) 
    
    Me.ComboBoxFilterColumn.Clear
    For Each cell In headerRange
        Me.ComboBoxFilterColumn.AddItem cell.Value
    Next cell
    
    PopulateListBox
End Sub

Private Sub btnFilter_Click()
    Dim selectedColumn As String
    Dim filterValue As String
    
    selectedColumn = Me.ComboBoxFilterColumn.Value
    filterValue = Me.TextBoxFilterValue.Value
    
    If selectedColumn = "" Then
        MsgBox "Please select a column to filter.", vbExclamation, "Filter Error"
        Exit Sub
    End If
    
    PopulateListBox selectedColumn, filterValue
End Sub

Private Sub btnClose_Click()

    Unload Me
End Sub

Private Sub PopulateListBox(Optional filterColumn As String = "", Optional filterValue As String = "")
    Dim ws As Worksheet
    Dim dataRange As Range
    Dim dataRow As Range
    Dim columnIndex As Integer
    Dim matchesCriteria As Boolean

    Set ws = ThisWorkbook.Sheets("Sheet1")
    Set dataRange = ws.Range("A2").CurrentRegion 
    
    Me.ListBoxSummary.Clear

    If filterColumn <> "" Then
        columnIndex = Application.Match(filterColumn, ws.Rows(1), 0)
    End If
    
    For Each dataRow In dataRange.Rows
        matchesCriteria = True
        If filterColumn <> "" And filterValue <> "" Then
            If dataRow.Cells(1, columnIndex).Value <> filterValue Then
                matchesCriteria = False
            End If
        End If
        
        If matchesCriteria Then
            Me.ListBoxSummary.AddItem Join(Application.Transpose(dataRow.Value), " | ")
        End If
    Next dataRow
End Sub
