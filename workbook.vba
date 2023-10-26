' Global Variables
public frontPage As Worksheet
public filterDatabase As Worksheet
Public productTable As ListObject
Public roomTable As ListObject
Public sortCell As Range
Public sortDirectionCell As Range
Public minColumnWidth As Integer

' Get Global Variables
Public Sub GetVariables()
    Set frontPage = ThisWorkbook.Worksheets(1)
    Set filterDatabase = ThisWorkbook.Worksheets(2)
    Set productTable = GetObject(1, "Product")
    Set roomTable = GetObject(2, "Room")
    Set sortCell = frontPage.Range("E3")
    Set sortDirectionCell = frontPage.Range("E4")
    minColumnWidth = 20
End Sub

' Get object from worksheet and return as ListObject
Public Function GetObject(Workbook As Integer, objectName As String) As ListObject
    Dim object As ListObject

    ' Attempt to find the table by it's name
    On Error Resume Next
    Set object = ThisWorkbook.Worksheets(Workbook).ListObjects(objectName)
    On Error GoTo 0

    ' Check ig the table was found. Give error if not, otherwise return ListObject
    If object Is Nothing Then
        ' Display a message that the object was not found
        MsgBox "The " + objectName + " object could not be found!"
        Exit Function
    Else
        Set GetObject = object
    End If
End Function

' Set the styles of the page and format the content
Public Sub SetPageStyle()
    Dim cell As Range
    Dim column As ListColumn
    Dim windowWidth As Double
    Dim columnCount As Integer
    Dim fontFactor As Double

    If Not productTable Is Nothing Then
        ' Set table stylings such as text wrapping and autofit
        productTable.Range.WrapText = False
        productTable.ShowAutoFilterDropDown = False
        productTable.Range.EntireColumn.AutoFit
        productTable.Range.EntireRow.AutoFit

        ' Change scroll limit of the page to fit the table
        frontPage.ScrollArea = "A:" + Split(Cells(1, productTable.ListColumns.Count).Address, "$")(1)

        ' Get the number of columns in the table
        columnCount = productTable.ListColumns.Count
        
        ' Calculate the font factor to account for the font and font size impact on cell width
        Set cell = Worksheets(1).Cells(1, 1)
        fontFactor = cell.Width / cell.ColumnWidth
        windowWidth = windowWidth / fontFactor
        
        ' Set autofit on all columns in the table, but then resize larger cells to a max width
        For Each column In productTable.ListColumns
            ' Set alignment of cells content
            column.Range.VerticalAlignment = xlCenter
            column.Range.HorizontalAlignment = xlCenter

            ' If column is too large, set it to max
            If column.DataBodyRange.ColumnWidth > minColumnWidth Then
                column.DataBodyRange.ColumnWidth = minColumnWidth
            End If
        Next column
        
        ' Resize the title to fit the new width
        Range(cell(1, 1), cell(1, columnCount)).Merge Across:=True
        
        ' Re-enable text wrapping
        productTable.Range.WrapText = True
    End If
End Sub

' Update product table with new rooms
Public Sub UpdateProductRooms()
    Dim rooms As Range
    Dim room As Range
    Dim col As ListColumn

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
    SetPageStyle
End Sub

' Sort & Filter options
Public Sub UpdateDropdowns()
    Dim sortOptions As String
    Dim column As ListColumn

    If Not productTable Is Nothing Then
        ' Remove Existing Sort Options
        sortCell.Validation.Delete
        
        ' Add each table heading to a string of options
        For Each column In productTable.ListColumns
            sortOptions = sortOptions + column.Name + ","
        Next

        ' Remove the trailing comma from the string of options
        sortOptions = Left(sortOptions, Len(sortOptions) - 1)

        ' Create the sort dropdown, and add a default value if the box is empty
        sortCell.Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:=sortOptions
        If sortCell.Value = "" Then
            sortCell.Value = "Default"
        End If
    End If
End Sub

' Sort product table
Public Sub SortProductTable()
    Dim sortDirection As Variant

    ' Identify correct sort order
    If sortDirectionCell.Value = "Descending" Then
        sortDirection = xlDescending
    Else
        sortDirection = xlAscending
    End If

    ' Apply sorting options to product table
    With productTable.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("Product[" + sortCell.Value + "]"), SortOn:=xlSortOnValues, Order:=sortDirection
        .Header = xlYes
        .Apply
    End With
End Sub

' Run at launch
Private Sub Workbook_Open()
    ' Get global variables
    GetVariables
End Sub


