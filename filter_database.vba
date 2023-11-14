' Run on change in sheet
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Update product rooms
    ThisWorkbook.UpdateProductRooms

    ' Update filter tables
    ThisWorkbook.UpdateDatabaseTables
End Sub

' Run when formula's are calculated
Private Sub Worksheet_Calculate()
    ' Update filter tables
    ThisWorkbook.UpdateDatabaseTables
End Sub