' Get object from worksheet and return as ListObject
Function getObject(workbook As Integer, objectName As String) As ListObject
    Dim object As ListObject

    ' Attempt to find the table by it's name
    On Error Resume Next
    Set object = Worksheets(workbook).ListObjects(objectName)
    On Error GoTo 0

    ' Check ig the table was found. Give error if not, otherwise return ListObject
    If object Is Nothing Then
        ' Display a message that the object was not found
        MsgBox "The " + objectName + " object could not be found!"
        Exit Function
    Else
        Set getObject = object
    End If
End Function

Function getConnection(connectionName As String)
    Dim connection As WorkbookConnection

    ' Attempt to find the connection by it's name
    On Error Resume Next
    Set connection = ThisWorkbook.Connections(connectionName)
    On Error GoTo 0

    ' Check if the connection was found
    If connection Is Nothing Then
        ' Display a message that the connection was not found
        MsgBox "The " + connectionName + " connection could not be found!"
        Exit Function
    Else
        Set getConnection = connection
    End If
End Function

' Resize results table to match usable window width
Sub ResizeResultTable()
    Dim resultTable As ListObject
    Dim cell As Range
    Dim column As ListColumn
    Dim windowWidth As Double
    Dim columnCount As Integer
    Dim fontFactor As Double

    Set resultTable = getObject(1, "Result")

    If Not resultTable Is Nothing Then
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
    
    Set conn = getConnection("Results_Connection")
    Set productTable = getObject(3, "Product")

    If Not conn Is Nothing And Not productTable Is Nothing Then
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
    
    Set roomTable = getObject(2, "Room")
    Set productTable = getObject(3, "Product")

    If Not roomTable Is Nothing And Not productTable Is Nothing Then
        ' Get rooms from table of rooms
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