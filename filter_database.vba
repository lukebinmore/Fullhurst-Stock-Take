' Run on change in sheet
Private Sub Worksheet_Change(ByVal Target As Range)
    ' Update product rooms
    ThisWorkbook.UpdateProductRooms
End Sub