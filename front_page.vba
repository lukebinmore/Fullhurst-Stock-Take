' Run when new room button clicked
Private Sub BTNRoom_Add_Click()
    ThisWorkbook.Wrap "AddNewRoom"

    ' Update page style
    ThisWorkbook.Wrap "SetPageStyle"
End Sub

' Run when new room button clicked
Private Sub BTNRoom_Delete_Click()
    ThisWorkbook.Wrap "DeleteRoom"

    ' Update page style
    ThisWorkbook.Wrap "SetPageStyle"
End Sub

' Run when create backup button clicked
Private Sub BTNBackup_Click()
    ThisWorkbook.Wrap "BackupDatabase"
End Sub

' Run when export data button clicked
Private Sub BTNExport_Click()
    ThisWorkbook.Wrap "ExportProductData"
End Sub

' Run when show/hide filters button clicked
Private Sub BTNHide_Filters_Click()
    ThisWorkbook.Wrap "ShowHideSection", "BTNHide_Filters"

    ' Update page style
    ThisWorkbook.Wrap "SetPageStyle"
End Sub

' Run when show/hide new product button clicked
Private Sub BTNHide_New_Click()
    ThisWorkbook.Wrap "ShowHideSection", "BTNHide_New"

    ' Update page style
    ThisWorkbook.Wrap "SetPageStyle"
End Sub

' Run when show/hide search button clicked
Private Sub BTNHide_Search_Click()
    ThisWorkbook.Wrap "ShowHideSection". "BTNHide_Search"

    ' Update page style
    ThisWorkbook.Wrap "SetPageStyle"
End Sub

' Run when add new product button clicked
Private Sub BTNNew_Click()
    ThisWorkbook.Wrap "AddNewProduct"

    ' Update databases
    ThisWorkbook.Wrap "UpdateDatabaseTables"

    ' Update page style
    ThisWorkbook.Wrap "SetPageStyle"
End Sub

' Run when reset filters button clicked
Private Sub BTNReset_Click()
    ThisWorkbook.Wrap "ResetFilters"

    ' Update page style
    ThisWorkbook.Wrap "SetPageStyle"
End Sub

' Run on change in sheet
Private Sub Worksheet_Change(ByVal Target As Range)
    With ThisWorkbook
        ' Get global variables in workbook
        .GetVariables

        ' Check what cell was targeted
        Select Case Target.Address
        Case .sortCell.Address, .sortDirectionCell.Address
            .Wrap "SortProductTable"
        Case .searchCell.Address, .searchFieldCell.Address
            .Wrap "SearchProductTable"
        End Select

        ' Check if cell triggered was in filter table
        If Not .filterTable Is Nothing Then
            If Not Intersect(.filterTable.DataBodyRange, Range(Target.Address)) Is Nothing Then
                .Wrap "SetProductFilters"
            End If
        End If

        ' Update sheet style
        .Wrap "SetPageStyle"
    End With
End Sub