' Run on change in sheet
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Get global variables in workbook
    ThisWorkbook.GetVariables

    ' Check what cell was targeted
    Select Case Target.Address
    Case ThisWorkbook.sortCell.Address, ThisWorkbook.sortDirectionCell.Address
        ThisWorkbook.SortProductTable
    End Select
End Sub