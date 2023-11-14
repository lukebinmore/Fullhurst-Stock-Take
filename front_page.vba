' Run when new room button clicked
Private Sub BTNRoom_Click()
    ThisWorkbook.AddNewRoom

    ' Update page style
    ThisWorkbook.SetPageStyle
End Sub

' Run when create backup button clicked
Private Sub BTNBackup_Click()
    ThisWorkbook.BackupDatabase
End Sub

' Run when export data button clicked
Private Sub BTNExport_Click()
    ThisWorkbook.ExportProductData
End Sub

' Run when show/hide filters button clicked
Private Sub BTNHide_Filters_Click()
    ThisWorkbook.ShowHideSection ("BTNHide_Filters")

    ' Update page style
    ThisWorkbook.SetPageStyle
End Sub

' Run when show/hide new product button clicked
Private Sub BTNHide_New_Click()
    ThisWorkbook.ShowHideSection ("BTNHide_New")

    ' Update page style
    ThisWorkbook.SetPageStyle
End Sub

' Run when show/hide search button clicked
Private Sub BTNHide_Search_Click()
    ThisWorkbook.ShowHideSection ("BTNHide_Search")

    ' Update page style
    ThisWorkbook.SetPageStyle
End Sub

' Run when add new product button clicked
Private Sub BTNNew_Click()
    ThisWorkbook.AddNewProduct

    ' Update databases
    ThisWorkbook.UpdateDatabaseTables

    ' Update page style
    ThisWorkbook.SetPageStyle
End Sub

' Run when reset filters button clicked
Private Sub BTNReset_Click()
    ThisWorkbook.ResetFilters

    ' Update page style
    ThisWorkbook.SetPageStyle
End Sub

' Run on change in sheet
Private Sub Worksheet_Change(ByVal Target As Range)
    With ThisWorkbook
        ' Get global variables in workbook
        .GetVariables

        ' Check what cell was targeted
        Select Case Target.Address
        Case .sortCell.Address, .sortDirectionCell.Address
            .SortProductTable
        Case .searchCell.Address, .searchFieldCell.Address
            .SearchProductTable
        End Select

        ' Check if cell triggered was in filter table
        If Not .filterTable Is Nothing Then
            If Not Intersect(.filterTable.DataBodyRange, Range(Target.Address)) Is Nothing Then
                .SetProductFilters
            End If
        End If

        ' Update sheet style
        .SetPageStyle
    End With
End Sub