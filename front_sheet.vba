'Resize results table to match usable window width
Sub ResizeResultsTable()
    Dim resultsTablee As ListObject
    Dim cell As Range
    Dim column As ListColumn
    Dim windowWidth As Double
    Dim columnCount As Integer
    Dim fontFactor As Double
    
    ' Set the target table
    Set resultsTable = Worksheets(1).ListObjects("Results")
    
    If resultsTable Is Nothing Then
        MsgBox "The 'Results' table does not exist on this sheet."
        Exit Sub
    End If
    
    ' Get the current window width in points
    windowWidth = Application.ActiveWindow.Width
    
    ' Get the number of columns in the table
    columnCount = resultsTable.ListColumns.Count
    
    ' Calculate the font factor to account for the font and font size impact on cell width
    Set cell = Worksheets(1).Cells(1, 1)
    fontFactor = cell.Width / cell.ColumnWidth
    
    ' Calculate the new width for each column based on the window width and font factor
    For Each column In resultsTable.ListColumns
        column.DataBodyRange.ColumnWidth = (windowWidth / columnCount) / fontFactor
    Next column
End Sub

'Refresh results table data
Sub RefreshResultsTable()
    Dim conn As WorkbookConnection
    
    ' Attempt to find the connection by its name
    On Error Resume Next
    Set conn = ThisWorkbook.Connections("Results_Connection")
    On Error GoTo 0
    
    ' Check if the connection was found
    If conn Is Nothing Then
        ' Display a message if the connection was not found
        MsgBox "Connection named 'Results_Connection' not found in this workbook."
    Else
        ' Refresh the connection if it was found
        conn.Refresh
    End If
End Sub



'Run at launch
Private Sub Workbook_Open()
    'Set Front Page scroll limit
    Worksheets(1).ScrollArea = "A:H"

    'Refresh table data
    RefreshResultsTable

    'Resize results table to fit window
    ResizeResultsTable
End Sub