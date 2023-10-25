'Resize results table to match usable window width
Sub ResizeResultTable()
    Dim resultTable As ListObject
    Dim cell As Range
    Dim column As ListColumn
    Dim windowWidth As Double
    Dim columnCount As Integer
    Dim fontFactor As Double
    
    ' Attempt to find the table by its name
    On Error Resume Next
    Set resultTable = Worksheets(1).ListObjects("Result")
    On Error GoTo 0
    
    ' Check if the tables were found
    If resultTable Is Nothing Then
        MsgBox "The 'Result' table does not exist on this sheet."
        Exit Sub
    Else
        ' Get the current window width in points
        windowWidth = Application.ActiveWindow.usableWidth
        
        ' Get the number of columns in the table
        columnCount = resultTable.ListColumns.Count
        
        ' Calculate the font factor to account for the font and font size impact on cell width
        Set cell = Worksheets(1).Cells(1, 1)
        fontFactor = cell.Width / cell.ColumnWidth
        
        ' Calculate the new width for each column based on the window width and font factor
        For Each column In resultTable.ListColumns
            column.DataBodyRange.ColumnWidth = (windowWidth / columnCount) / fontFactor
        Next column
        
        ' Resize the title to fit the new width
        Range(cell(1, 1), cell(1, columnCount)).Merge Across:=True
    End If
End Sub

'Refresh results table data
Sub RefreshResultTable()
    Dim conn As WorkbookConnection
    Dim productTable As ListObject
    
    ' Attempt to find the connection by its name
    On Error Resume Next
    Set conn = ThisWorkbook.Connections("Results_Connection")
    Set productTable = Worksheets(3).ListObjects("Product")
    On Error GoTo 0
    
    ' Check if the connection and table were found
    If conn Is Nothing Then
        ' Display a message if the connection was not found
        MsgBox "Connection named 'Results_Connection' not found in this workbook."
    ElseIf productTable Is Nothing Then
        ' Display a message if the table was not found
        MsgBox "The 'Product' table does not exist on this sheet."
        Exit Sub
    Else
        ' Refresh the page cell width limit
        Worksheets(1).ScrollArea = "A:" + Split(Cells(1, productTable.ListColumns.Count).Address, "$")(1)
        
        ' Refresh the connection
        conn.Refresh
    End If
End Sub

'Update product table with new rooms
Sub UpdateProductRooms()
    Dim roomTable As ListObject
    Dim productTable As ListObject
    Dim rooms As Range
    Dim room As Range
    Dim col As ListColumn
    
    'Attempt to find the tables by name
    On Error Resume Next
    Set roomTable = Worksheets(2).ListObjects("Room")
    Set productTable = Worksheets(3).ListObjects("Product")
    On Error GoTo 0
    
    ' Check if the tables were found
    If roomTable Is Nothing Then
        ' Display a message if the table was not found
        MsgBox "The 'Room' table does not exist on this sheet."
        Exit Sub
    ElseIf productTable Is Nothing Then
        ' Display a message if the table was not found
        MsgBox "The 'product' table does not exist on this sheet."
        Exit Sub
    Else
        ' Get the list of rooms that should now exist
        Set rooms = roomTable.ListColumns(1).DataBodyRange

        For Each room In rooms
            ' Check if the current cell is empty
            If Not Trim(room.Value) = "" Then
                On Error Resume Next
                ' Reset the col object & attempt to find new room column
                Set col = Nothing
                Set col = productTable.ListColumns(room.Value)
                On Error GoTo 0
                ' Check if the room column exists in the product table
                If col Is Nothing Then
                    ' Add new room column
                    productTable.ListColumns.Add.Name = room.Value
                End If
            End If
        Next
    End If
    
    ' Refresh and resize the results table
    RefreshResultTable
    ResizeResultTable
End Sub

'Run at launch
Private Sub Workbook_Open()
    'Refresh table data
    RefreshResultTable

    'Resize results table to fit window
    ResizeResultTable
End Sub