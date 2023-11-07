' Global Variables
Public frontPage As Worksheet
Public filterDatabase As Worksheet
Public productTable As ListObject
Public filterTable As ListObject
Public campusTable As ListObject
Public typeTable As ListObject
Public supplierTable As ListObject
Public roomTable As ListObject
Public subjectTable As ListObject
Public newProductTable As ListObject
Public sortCell As Range
Public searchCell As Range
Public sortDirectionCell As Range
Public minColumnWidth As Integer
Public maxColumnWidth As Integer

' Get Global Variables
Public Sub GetVariables()
    Set frontPage = ThisWorkbook.Worksheets(1)
    Set filterDatabase = ThisWorkbook.Worksheets(2)
    Set productTable = GetObject(1, "Product")
    Set filterTable = GetObject(1, "Filter")
    Set campusTable = GetObject(2, "Campus")
    Set typeTable = GetObject(2, "Type")
    Set supplierTable = GetObject(2, "Supplier")
    Set roomTable = GetObject(2, "Room")
    Set subjectTable = GetObject(2, "Subject")
    Set newProductTable = GetObject(1, "NewProduct")
    Set sortCell = frontPage.Range("E3")
    Set sortDirectionCell = frontPage.Range("E4")
    Set searchCell = frontPage.Range("B14")
    minColumnWidth = 3
    maxColumnWidth = 40
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

' Check value exists in table
Public Function CheckValueExists(ByVal value As String, ByVal table As ListObject)
    If IsError(Application.Match(value, table.ListColumns(1).DataBodyRange, 0)) Then
        CheckValueExists = False
    Else
        CheckValueExists = True
    End If
End Function

' Set the styles of the page and format the content
Public Sub SetPageStyle()
    ' Disable events & screen updating
    SetScreenEvents(False)
    
    Dim cell As Range
    Dim column As ListColumn
    Dim columnCount As Integer

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Get global variables
    GetVariables

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
        
        ' Set autofit on all columns in the table, but then resize larger cells to a max width
        For Each column In productTable.ListColumns
            ' Set alignment of cells content
            column.Range.VerticalAlignment = xlCenter
            column.Range.HorizontalAlignment = xlCenter
            
            ' If column is too large, set it to max
            If column.Range.ColumnWidth < minColumnWidth Then
                column.Range.ColumnWidth = minColumnWidth
            ElseIf column.Range.ColumnWidth > maxColumnWidth Then
                column.Range.ColumnWidth = maxColumnWidth
            End If
        Next column
        
        ' Resize the title to fit the new width
        Set cell = Worksheets(1).Cells(1, 1)
        Range(cell(1, 1), cell(1, columnCount)).Merge Across:=True
        
        ' Re-enable text wrapping
        productTable.Range.WrapText = True
    End If

    ' Provide error message to user
ErrorHandler:
        ErrorMessage

    ' Re-Enable events & screen updating
    SetScreenEvents(True)
End Sub

' Update product table with new rooms
Public Sub UpdateProductRooms()
    ' Disable events & screen updating
    SetScreenEvents(False)
    
    Dim rooms As Range
    Dim room As Range
    Dim col As ListColumn

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Get global variables
    GetVariables

    If Not roomTable Is Nothing And Not productTable Is Nothing Then
        ' Get rooms from table of rooms
        Set rooms = roomTable.ListColumns(1).DataBodyRange
        
        ' Loop through each room in the table
        For Each room In rooms
            ' Check if the current cell is empty
            If Not Trim(room.value) = "" Then
                ' Reset the col object & attempt to find new room column
                Set col = Nothing
                On Error Resume Next
                Set col = productTable.ListColumns(room.value)

                ' Check if the room column exists in the product table
                If col Is Nothing Then
                    ' Change scroll limit of the page to fit the table
                    frontPage.ScrollArea = "A:" + Split(Cells(1, productTable.ListColumns.Count + 1).Address, "$")(1)

                    ' Add new room column
                    productTable.ListColumns.Add.Name = room.value
                End If
            End If
        Next
    End If
    
    ' Resize the product table
    SetPageStyle

    ' Provide error message to user
ErrorHandler:
        ErrorMessage

    ' Re-Enable events & screen updating
    SetScreenEvents(True)
End Sub

' Sort & Filter options
Public Sub UpdateDropdowns()
    ' Disable events & screen updating
    SetScreenEvents(False)
    
    Dim sortOptions As String
    Dim column As ListColumn

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Get global variables
    GetVariables

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
        If sortCell.value = "" Then
            sortCell.value = "Default"
        End If
    End If

    ' Set page style
    SetPageStyle

    ' Provide error message to user
ErrorHandler:
        ErrorMessage

    ' Re-Enable events & screen updating
    SetScreenEvents(True)
End Sub

' Sort product table
Public Sub SortProductTable()
    ' Disable events & screen updating
    SetScreenEvents(False)
    
    Dim sortDirection As Variant

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Get global variables
    GetVariables

    ' Identify correct sort order
    If sortDirectionCell.value = "Descending" Then
        sortDirection = xlDescending
    Else
        sortDirection = xlAscending
    End If

    ' Apply sorting options to product table
    With productTable.Sort
        .SortFields.Clear
        .SortFields.Add Key:=Range("Product[" + sortCell.value + "]"), SortOn:=xlSortOnValues, Order:=sortDirection
        .Header = xlYes
        .Apply
    End With

    ' Set page style
    SetPageStyle

    ' Provide error message to user
ErrorHandler:
        ErrorMessage

    ' Re-Enable events & screen updating
    SetScreenEvents(True)
End Sub

' Search the product table with text
Public Sub SearchProductTable()
    ' Disable events & screen updating
    SetScreenEvents(False)

    Dim searchInput As String

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Get global variables
    GetVariables
    
    ' Clear previous search filter
    productTable.Range.AutoFilter Field:=1

    ' Filter the table with the value in the cell
    If Not searchCell.value = "" Then
        productTable.Range.AutoFilter Field:=1, Criteria1:= _
        "=*" & searchCell.value & "*", Operator:=xlAnd
    End If

    ' Set page style
    SetPageStyle

    ' Provide error message to user
ErrorHandler:
        ErrorMessage

    ' Re-Enable events & screen updating
    SetScreenEvents(True)
End Sub

' Add new items to product table
Public Sub AddNewProduct()
    ' Disable events & screen updating
    SetScreenEvents(False)

    Dim newName, newDesc, newType, newSupplier, newProdCode, newSubject, newCampus, newRoom As String
    Dim newQuantity As Integer
    Dim newRow As ListRow
    Dim roomRowIndex As Integer

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Get Variables
    GetVariables

    ' Get data from table
    With newProductTable.DataBodyRange
        newName = .Cells(1, 1)
        newDesc = .Cells(1, 2)
        newType = .Cells(1, 3)
        newSupplier = .Cells(1, 4)
        newProdCode = .Cells(1, 5)
        newSubject = .Cells(1, 6)
        newCampus = .Cells(1, 7)
        newRoom = .Cells(1, 8)
        newQuantity = .Cells(1, 9)
    End With

    ' Check if type exists, add it to table if it doesn't
    If Not CheckValueExists(newType, typeTable) And Not newType = "" Then
        Dim typeRow As ListRow
        Set typeRow = typeTable.ListRows.Add
        typeRow.Range(1) = newType
    End If

    ' Check if supplier exists, add it to table if it doesn't
    If Not CheckValueExists(newSupplier, supplierTable) And Not newSupplier = "" Then
        Dim supplierRow As ListRow
        Set supplierRow = supplierTable.ListRows.Add
        supplierRow.Range(1) = newSupplier
    End If

    ' Check if subject exists, add it to table if it doesn't
    If Not CheckValueExists(newSubject, subjectTable) And Not newSubject = "" Then
        Dim subjectRow As ListRow
        Set subjectRow = subjectTable.ListRows.Add
        subjectRow.Range(1) = newSubject
    End If

    ' Check if campus exists, add it to table if it doesn't
    If Not CheckValueExists(newCampus, campusTable) And Not newCampus = "" Then
        Dim campusRow As ListRow
        Set campusRow = campusTable.ListRows.Add
        campusRow.Range(1) = newCampus
    End If

    ' Check if room exists, add it to table if it doesn't
    If Not CheckValueExists(newRoom, roomTable) And Not newRoom = "" Then
        Dim roomRow As ListRow
        Set roomRow = roomTable.ListRows.Add
        roomRow.Range(1) = newRoom

        ' Update product table with new room
        UpdateProductRooms
    End If

    ' Set row index of room column in product table
    roomRowIndex = productTable.ListColumns(newRoom).Index

    ' Add new row
    Set newRow = productTable.ListRows.Add
    With newRow
        .Range(1) = newName
        .Range(2) = newDesc
        .Range(3) = newType
        .Range(4) = newSupplier
        .Range(5) = newProdCode
        .Range(6) = newSubject
        .Range(7) = newCampus

        ' Check if room has been entered
        If Not newRoom = "" Then
            .Range(roomRowIndex) = newQuantity
        End If
    End With

    ' Update dropdowns and refresh styling
    UpdateDropdowns
    SetPageStyle

    ' Provide error message to user
ErrorHandler:
        ErrorMessage

    ' Re-Enable events & screen updating
    SetScreenEvents(True)
End Sub

' Display error message
Public Sub ErrorMessage()
    If Err.Number <> 0 Then
        MsgBox "Something went wrong, Please try again"
    End If
End Sub

' Reset filters
Public Sub ResetFilters()
    ' Disable events & screen updating
    SetScreenEvents(False)

    Dim column As ListColumn

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Get global variables
    GetVariables

    If Not filterTable Is Nothing Then
        ' Clear all filter values
        For Each column In filterTable.ListColumns
            filterTable.DataBodyRange.Cells(1, column.Index).value = ""
            filterTable.DataBodyRange.Cells(2, column.Index).value = ""
        Next
    End If

    ' Apply new filters to product table
    ApplyProductFilters

    ' Provide error message to user
ErrorHandler:
        ErrorMessage

    ' Re-Enable events & screen updating
    SetScreenEvents(True)
End Sub

' Filter table base on user choices
Public Sub SetProductFilters()
    ' Disable events & screen updating
    SetScreenEvents(False)

    Dim column As ListColumn

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Get global variables
    GetVariables

    ' Complete checks for each column in filter table
    For Each column In filterTable.ListColumns
        Dim appliedFilterCell As Range
        Dim newFilter As Range

        ' Get data from column input and current filters
        Set appliedFilterCell = filterTable.DataBodyRange.Cells(2, column.Index)
        Set newFilterCell = filterTable.DataBodyRange.Cells(1, column.Index)

        ' Check if value has been entered into input cell
        If Not newFilterCell.value = "" Then
            Dim filter As Variant
            Dim exists As Boolean
            
            ' Assume that the new filter doesn't already exist
            exists = False

            ' Check if item already exists in applied filters
            For Each filter In Split(appliedFilterCell.value, ",")
                If filter = newFilterCell.value Then
                    exists = True
                End If
            Next

            ' If item is new, add it to filter list
            If Not exists Then
                If Not appliedFilterCell.value = "" Then
                    appliedFilterCell.value = appliedFilterCell.value + "," + newFilterCell.value
                Else
                    appliedFilterCell.value = newFilterCell.value
                End If
            End If

            ' Clear input cell
            newFilterCell.value = ""
        End If
    Next

    ' Apply new filters
    ApplyProductFilters

    ' Provide error message to user
ErrorHandler:
        ErrorMessage

    ' Re-Enable events & screen updating
    SetScreenEvents(True)
End Sub

' Apply product table filters
Public Sub ApplyProductFilters()
    ' Disable events & screen updating
    SetScreenEvents(False)

    Dim column As ListColumn

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Get global variables
    GetVariables

    ' Complete checks for each column in filter table
    For Each column In filterTable.ListColumns
        Dim filtersCell As Range
        Dim filters() As String
        Dim filter As Variant
        Dim productColumnIndex As Integer

        ' Get filters to apply
        Set filtersCell = filterTable.DataBodyRange.Cells(2, column.Index)
        filters = Split(filtersCell.value, ",")

        ' Get product column index to apply filter to
        productColumnIndex = productTable.ListColumns(column.Name).Index

        ' Clear previous filters
        productTable.Range.AutoFilter Field:=productColumnIndex

        ' Filter the table with the value in the cell
        If Not filtersCell.value = "" Then
            productTable.Range.AutoFilter Field:=productColumnIndex, Criteria1:=filters, Operator:=xlFilterValues
        End If
    Next

    ' Set page style
    SetPageStyle

    ' Provide error message to user
ErrorHandler:
        ErrorMessage

    ' Re-Enable events & screen updating
    SetScreenEvents(True)
End Sub

' Export filtered product data to new file
Public Sub ExportProductData()
    ' Disable events & screen updating
    SetScreenEvents(False)

    Dim newWorkbook As Workbook
    Dim savePath As String
    Dim saveFileName As String

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Get global variables
    GetVariables

    ' Check if table exists
    If Not productTable Is Nothing Then
        ' Prompt the user for a save location and file name
        savePath = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx")
        
        ' Check if the user canceled the save dialog
        If savePath <> "False" Then
            Dim column As ListColumn

            ' Create a new workbook
            Set newWorkbook = Workbooks.Add

            ' Copy filtered data to new workbook
            For Each column In productTable.ListColumns
                column.Range.SpecialCells(xlCellTypeVisible).Copy Destination:=newWorkbook.Worksheets(1).Cells(1, column.Index)
            Next

            ' Set new workbook stylings
            With newWorkbook.Worksheets(1)
                .PageSetup.Orientation = xlLandscape
                .Cells.WrapText = False
                .Cells.EntireRow.AutoFit
                .Cells.EntireColumn.AutoFit
            End With

            ' Save the new workbook with the user-specified name
            saveFileName = Mid(savePath, InStrRev(savePath, "\") + 1)
            newWorkbook.SaveAs savePath
            newWorkbook.Close SaveChanges:=False
        End If
    End If

ErrorHandler:
    ErrorMessage

    ' Re-Enable events & screen updating
    SetScreenEvents(True)
End Sub

' Enable or disable screen updating and event catching
Public Sub SetScreenEvents(ByVal state As Boolean)
    Application.EnableEvents = state
    Application.ScreenUpdating = state
End Sub


' Run at launch
Private Sub Workbook_Open()
    ' Ensure events and screen updating are enabled
    SetScreenEvents(True)

    ' Apply page style
    SetPageStyle
End Sub