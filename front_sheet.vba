' Get object from worksheet and return as ListObject
Function getObject(Workbook As Integer, objectName As String) As ListObject
    Dim object As ListObject

    ' Attempt to find the table by it's name
    On Error Resume Next
    Set object = Worksheets(Workbook).ListObjects(objectName)
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

' Resize the Product table to match usable window width
Sub ResizeProductTable()
    Dim productTable As ListObject
    Dim cell As Range
    Dim column As ListColumn
    Dim windowWidth As Double
    Dim columnCount As Integer
    Dim fontFactor As Double

    Set productTable = getObject(1, "Product")

    If Not productTable Is Nothing Then
        ' Disable text wrapping while calculating widths
        productTable.Range.WrapText = False

        ' Change scroll limit of the page to fit the table
        Worksheets(1).ScrollArea = "A:" + Split(Cells(1, productTable.ListColumns.Count).Address, "$")(1)

        ' Get the current window width in points, and the number of columns in the table
        windowWidth = Application.ActiveWindow.usableWidth
        columnCount = productTable.ListColumns.Count
        
        ' Calculate the font factor to account for the font and font size impact on cell width
        Set cell = Worksheets(1).Cells(1, 1)
        fontFactor = cell.Width / cell.ColumnWidth
        
        ' Calculate the new width for each column based on the window width and font factor
        For Each column In productTable.ListColumns
            column.DataBodyRange.ColumnWidth = (windowWidth / columnCount) / fontFactor
        Next column
        
        ' Resize the title to fit the new width
        Range(cell(1, 1), cell(1, columnCount)).Merge Across:=True
        
        ' Re-enable text wrapping
        productTable.Range.WrapText = True
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
    Set productTable = getObject(1, "Product")

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
    
    ' Resize the product table
    ResizeProductTable
End Sub

'Run at launch
Private Sub Workbook_Open()
    'Resize the product table to fit window
    ResizeProductTable
End Sub
