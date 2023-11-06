' Run when add new product button clicked
Private Sub BTNNew_Click()
    ThisWorkbook.AddNewProduct
End Sub

' Run when reset filters button clicked
Private Sub BTNReset_Click()
    ThisWorkbook.ResetFilters
End Sub

' Run on change in sheet
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Get global variables in workbook
    ThisWorkbook.GetVariables

    ' Check what cell was targeted
    Select Case Target.Address
    Case ThisWorkbook.sortCell.Address, ThisWorkbook.sortDirectionCell.Address
        ThisWorkbook.SortProductTable
    Case ThisWorkbook.searchCell.Address
        ThisWorkbook.SearchProductTable
    End Select

    ' Check if cell triggered was in filter table
    If Not ThisWorkbook.filterTable Is Nothing Then
        If Not Intersect(ThisWorkbook.filterTable.DataBodyRange, Range(Target.Address)) Is Nothing Then
            ThisWorkbook.SetProductFilters
        End If
    End If
End Sub