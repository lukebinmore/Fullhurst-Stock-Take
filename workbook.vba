' Global Variables
Public frontPage As Worksheet
Public filterDatabase As Worksheet
Public productTable As ListObject
Public filterTable As ListObject
Public roomTable As ListObject
Public newProductTable As ListObject
Public sortCell As Range
Public searchCell As Range
Public searchFieldCell As Range
Public sortDirectionCell As Range
Public filterSection As Range
Public newProductSection As Range
Public searchSection As Range
Public resetButtonCell As Range
Public addButtonCell As Range
Public minColumnWidth As Integer
Public maxColumnWidth As Integer

' Get Global Variables
Public Sub GetVariables()
    Set frontPage = ThisWorkbook.Worksheets(1)
    Set filterDatabase = ThisWorkbook.Worksheets(2)
    Set productTable = GetObject(1, "Product")
    Set filterTable = GetObject(1, "Filter")
    Set roomTable = GetObject(2, "Room")
    Set newProductTable = GetObject(1, "NewProduct")
    Set sortCell = frontPage.Range("C5")
    Set sortDirectionCell = frontPage.Range("D5")
    Set searchCell = frontPage.Range("B14")
    Set searchFieldCell = frontPage.Range("H14")
    Set filterSection = frontPage.Range("A4:A8").EntireRow
    Set newProductSection = frontPage.Range("A9:A12").EntireRow
    Set searchSection = frontPage.Range("A13:A14").EntireRow
    Set resetButtonCell = frontPage.Range("F7")
    Set addButtonCell = frontPage.Range("H11")
    minColumnWidth = 3
    maxColumnWidth = 40
End Sub

' Get object from worksheet and return as ListObject
Public Function GetObject(workbook As Integer, objectName As String) As ListObject
    Dim object As ListObject

    ' Attempt to find the table by it's name
    On Error Resume Next
    Set object = ThisWorkbook.Worksheets(workbook).ListObjects(objectName)
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
    Dim column As ListColumn
    Dim columnCount As Integer

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Set table stylings such as text wrapping and autofit
    With frontPage.Cells.SpecialCells(xlCellTypeVisible)
        .WrapText = False
        .EntireColumn.AutoFit
        .EntireRow.AutoFit
        productTable.ShowAutoFilterDropDown = False
    End With

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
        If Not column.Range.EntireColumn.Hidden = True Then
            With column.Range
                If .ColumnWidth < minColumnWidth Then
                    .ColumnWidth = minColumnWidth
                ElseIf .ColumnWidth > maxColumnWidth Then
                    .ColumnWidth = maxColumnWidth
                End If
            End With
        End If
    Next column
    
    ' Re-enable text wrapping
    productTable.Range.WrapText = True

    Exit Sub

    ' Provide error message to user
ErrorHandler:
        ErrorMessage
End Sub

' Update database tables
Public Sub UpdateDatabaseTables()
    Dim table As ListObject

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Loop through tables in database sheet
    For Each table In filterDatabase.ListObjects
        ' Skip Room table
        If Not table.Name = "Room" Then
            Dim cell As Range
            Dim lastCell As Range
            Dim emptyCellsCleared As Boolean

            ' Loop through each row in the table
            For i = table.ListRows.Count To 1 Step -1
                ' Get last cell in table column
                Set lastCell = table.DataBodyRange.Cells(table.ListRows.Count, 1)

                ' Add row if not exmpty cell in last row, delete other empty cells
                With table.DataBodyRange.Cells(i, 1)
                    If Not .Value = "" And lastCell.Address = .Address Then
                        table.ListRows.Add
                    ElseIf .Value = "" Then
                        table.ListRows(i).Delete
                    End If
                End With
            Next

            ' Insert single cell if table is empty
            If table.ListRows.Count = 0 Then
                table.ListRows.Add
            End If
        End If
    Next

    Exit Sub

    ' Provide error message to user
ErrorHandler:
    ErrorMessage
End Sub

' Update product table with new rooms
Public Sub UpdateProductRooms()    
    Dim rooms As Range
    Dim room As Range
    Dim col As ListColumn

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Get rooms from table of rooms
    Set rooms = roomTable.ListColumns(1).DataBodyRange

    ' Exit early if no rooms have been entered yet
    If rooms Is Nothing Then
        Exit Sub
    End If
    
    ' Loop through each room in the table
    For Each room In rooms
        ' Check if the current cell is empty
        If Not room.value = "" Then
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

    Exit Sub

    ' Provide error message to user
ErrorHandler:
        ErrorMessage
End Sub

' Sort product table
Public Sub SortProductTable()    
    Dim sortDirection As Variant

    ' Enable error handeling
    On Error GoTo ErrorHandler

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

    Exit Sub

    ' Provide error message to user
ErrorHandler:
        ErrorMessage
End Sub

' Search the product table with text
Public Sub SearchProductTable()
    Dim searchColumn As Integer

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Check if search field is empty, set to default if it is
    If searchFieldCell.Value = "" Then
        searchFieldCell.Value = "Name"
    End If

    ' Set column to filter
    searchColumn = productTable.ListColumns(searchFieldCell.Value).Index
    
    ' Clear previous search filter
    productTable.Range.AutoFilter Field:=productTable.ListColumns("Name").Index
    productTable.Range.AutoFilter Field:=productTable.ListColumns("Description").Index
    productTable.Range.AutoFilter Field:=productTable.ListColumns("Product Code").Index

    ' Filter the table with the value in the cell
    If Not searchCell.value = "" Then
        productTable.Range.AutoFilter Field:=searchColumn, Criteria1:= _
        "=*" & searchCell.value & "*", Operator:=xlAnd
    End If

    Exit Sub

    ' Provide error message to user
ErrorHandler:
        ErrorMessage
End Sub

' Add new items to product table
Public Sub AddNewProduct()
    Dim newRow As ListRow
    Dim column As ListColumn

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Add new row
    If productTable.ListRows.Count = 0 Then
        Set newRow = productTable.ListRows.Add
    Else
        Set newRow = productTable.ListRows.Add(1)
    End If

    ' Insert data into product table
    For i = 1 To newProductTable.ListColumns.Count
        newRow.Range(i) = newProductTable.DataBodyRange.Cells(1, i)
    Next

    Exit Sub

    ' Provide error message to user
ErrorHandler:
        ErrorMessage
End Sub

' Display error message
Public Sub ErrorMessage()
    DIm errorMessage As String
    If Err.Number <> 0 Then
        errorMessage = "Something went wrong, Please try again." & vbCrlF & vbCrLf
        errorMessage = errorMessage & "Error: " & Err.Number & vbCrLf & vbCrLf
        errorMessage = errorMessage & Err.Description
        MsgBox errorMessage
    End If
End Sub

' Reset filters
Public Sub ResetFilters()
    Dim column As ListColumn

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Clear all filter values
    For Each column In filterTable.ListColumns
        filterTable.DataBodyRange.Cells(1, column.Index).value = ""
        filterTable.DataBodyRange.Cells(2, column.Index).value = ""
    Next

    ' Apply new filters to product table
    ApplyProductFilters

    Exit Sub

    ' Provide error message to user
ErrorHandler:
        ErrorMessage
End Sub

' Filter table base on user choices
Public Sub SetProductFilters()
    Dim column As ListColumn

    ' Enable error handeling
    On Error GoTo ErrorHandler

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

    Exit Sub

    ' Provide error message to user
ErrorHandler:
        ErrorMessage
End Sub

' Apply product table filters
Public Sub ApplyProductFilters()
    Dim column As ListColumn

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Complete checks for each column in filter table
    For Each column In filterTable.ListColumns
        Dim filtersCell As Range
        Dim filters() As String
        Dim room As Range
        Dim productColumnIndex As Integer

        ' Get filters to apply
        Set filtersCell = filterTable.DataBodyRange.Cells(2, column.Index)
        filters = Split(filtersCell.value, ",")

        ' Check if filtering rooms
        If column.Name = "Room" Then
            ' Show all columns
            Columns.Entirecolumn.Hidden = False

            ' Hide the columns based on the value of the cell
            If Not filtersCell.Value = "" Then
                For Each room In roomTable.ListColumns(1).DataBodyRange
                    If Not (UBound(Filter(filters, room.Value)) > -1) Then
                        productTable.ListColumns(room.Value).Range.EntireColumn.Hidden = True
                    End If
                Next
            End If
        Else
            ' Get product column index to apply filter to
            productColumnIndex = productTable.ListColumns(column.Name).Index

            ' Clear previous filters
            productTable.Range.AutoFilter Field:=productColumnIndex

            ' Filter the table with the value in the cell
            If Not filtersCell.Value = "" Then
                productTable.Range.AutoFilter Field:=productColumnIndex, Criteria1:=filters, Operator:=xlFilterValues
            End If
        End If
    Next

    Exit Sub

    ' Provide error message to user
ErrorHandler:
        ErrorMessage
End Sub

' Export filtered product data to new file
Public Sub ExportProductData()
    Dim newWorkbook As Workbook
    Dim savePath As String
    Dim saveFileName As String

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Prompt the user for a save location and file name
    savePath = Application.GetSaveAsFilename(FileFilter:="Excel Files (*.xlsx), *.xlsx")
    
    ' Check if the user canceled the save dialog
    If savePath <> "False" Then
        Dim column As ListColumn
        Dim newColumnIndex As Integer

        ' Create a new workbook
        Set newWorkbook = Workbooks.Add
        newColumnIndex = 1

        ' Copy filtered data to new workbook, use custom index to account for hidden columns
        For Each column In productTable.ListColumns
            If column.Range.EntireColumn.Hidden = False Then
                column.Range.SpecialCells(xlCellTypeVisible).Copy Destination:=newWorkbook.Worksheets(1).Cells(1, newColumnIndex)
                newColumnIndex = newColumnIndex + 1
            End If
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

    Exit Sub

ErrorHandler:
    ErrorMessage
End Sub

' Backup database to new file
Public Sub BackupDatabase()
    Dim defaultFileName, savePath, dateNow As String

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Get the current date in YYYY-MM-DD format
    dateNow = Format(Now, "YYYY-MM-DD")

    ' Contruct default file name
    defaultFileName = "Inventory Tracker Backup " & dateNow

    ' Get save directory and confirm file name from user
    savePath = Application.GetSaveAsFilename(InitialFileName:=defaultFileName, FileFilter:="Excel Macro-Enabled Workbook (*.xlsm), *.xlsm")

    ' Save the file if a path was entered
    If savePath <> "False" Then
        ThisWorkbook.SaveCopyAs Filename:=savePath
    End If

    Exit Sub

ErrorHandler:
    ErrorMessage
End Sub

' Enable or disable screen updating and event catching
Public Sub SetScreenEvents(ByVal state As Boolean)
    Application.EnableEvents = state
    Application.ScreenUpdating = state
End Sub

' Hide or show sections
Public Sub ShowHideSection(ByVal target As String)
    Dim button As Variant
    Dim resetButton As Variant
    DIm addButton As Variant
    Dim section As Range
    Dim buttonText() As String

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Collect correct section
    Select Case target
        Case "BTNHide_Filters"
            Set section = filterSection
        Case "BTNHide_New"
            Set section = newProductSection
        Case "BTNHide_Search"
            Set section = searchSection
    End Select

    ' Set buttons
    Set button = frontPage.OLEObjects(target).Object
    Set resetButton = frontPage.OLEObjects("BTNReset")
    Set addButton = frontPage.OLEObjects("BTNNew")
    buttonText = Split(button.Caption, " ")

    ' Change button text and set hidden state
    If section.Hidden = True Then
        section.Hidden = False
        button.Caption = "Hide " + buttonText(1)
    Else
        section.Hidden = True
        button.Caption = "Show " + buttonText(1)
    End If

    ' Fix button placement
    With resetButton
        .Visible = True
        If filterSection.Hidden = True Then
            .Visible = False
        End If
        .Top = resetButtonCell.Top
        .Left = resetButtonCell.Left
    End WIth
    With addButton
        .Visible = True
        If newProductSection.Hidden = True Then
            .Visible = False
        End If
        .Top = addButtonCell.Top
        .Left = addButtonCell.Left
    End With

    Exit Sub

ErrorHandler:
    ErrorMessage
End Sub

' Add new rooms to the product table
Public Sub AddNewRoom()
    Dim newRow As ListRow
    Dim newRoomName As String

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Get new room name from user
    newRoomName = InputBox("Please enter the new room number or name. E.G. G101", "Add New Room")

    ' Check if user entered a name
    If Not newRoomName = "" Then
        ' Add new room to room table
        Set newRow = roomTable.ListRows.Add
        newRow.Range(1) = newRoomName

        ' Update product table on frontPage
        UpdateProductRooms
    End If

    ' Re-apply existing table style
    productTable.TableStyle = productTable.TableStyle

    Exit Sub

ErrorHandler:
    ErrorMessage
End Sub

' Add new rooms to the product table
Public Sub DeleteRoom()
    Dim roomName, confirmation As String
    DIm rowIndex As Integer
    Dim col As ListColumn

    ' Enable error handeling
    On Error GoTo ErrorHandler

    ' Get room name from user
    roomName = InputBox("Please enter the room name you would like to delete. E.G. G101", "Delete Room")

    ' Check if user entered a name
    If Not roomName = "" Then
        ' Find room row in room table
        rowIndex = 0
        For i = roomTable.ListRows.Count To 1 Step -1
            If roomName = roomTable.DataBodyRange.Cells(i, 1).Value Then
                rowIndex = i
            End If
        Next

        ' Find room in product table
        On Error Resume Next
        Set col = productTable.ListColumns(roomName)
        On Error GoTo ErrorHandler

        ' Check if room exists in either table
        If col Is Nothing And rowIndex = 0 Then
            MsgBox "Error: Room not found!"
            Exit Sub
        End If

        ' Confirm user wishes to delete room
        confirmation = MsgBox("Are you sure you wish to delete " & roomName & "?",vbQuestion + vbYesNo, "Delete Room")
        
        ' Delete room if confirmed
        If confirmation = vbYes Then
            ' Delete from product table
            If Not col Is Nothing Then
                productTable.ListColumns(roomName).Delete
            End If
            
            ' Delete from room table
            If Not rowIndex = 0 Then
                roomTable.ListRows(rowIndex).Delete
            End If
        End If
    End If

    Exit Sub

ErrorHandler:
    ErrorMessage
End Sub

' Wrapper sub to disbale screen updating
Public Sub Wrap(subName As String, ParamArray args() As Variant)
    ' Disable events and screen updating
    SetScreenEvents(False)

    ' Check if variables are available
    If frontPage Is Nothing Then
        GetVariables
    End If

    ' Run provided sub or function with or without arguments
    If IsMissing(args) Then
        Application.Run "ThisWorkbook." & subName
    Else
        Application.Run "ThisWorkbook." & subName, args(0)
    End If

    ' Disable events and screen updating
    SetScreenEvents(True)
End Sub

' Run at launch
Private Sub Workbook_Open()
    ' Disable events and screen udpating
    SetScreenEvents(false)

    ' Get global variables
    GetVariables

    ' Update database tables
    UpdateDatabaseTables

    ' Sort product table
    SortProductTable

    ' Apply page style
    SetPageStyle

    ' Enable events and screen updating
    SetScreenEvents(True)
End Sub