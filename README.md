# **Inventory Tracker Usage Guide**

- [**Inventory Tracker Usage Guide**](#inventory-tracker-usage-guide)
  - [Overview](#overview)
  - [Layout](#layout)
    - [Management Buttons](#management-buttons)
    - [Filtering Section](#filtering-section)
    - [New Item Section](#new-item-section)
    - [Search Section](#search-section)
    - [Data Section](#data-section)
  - [Features](#features)
    - [Adding New Items](#adding-new-items)
    - [Editing Existing Items](#editing-existing-items)
    - [Deleting Items](#deleting-items)
    - [Searching Items](#searching-items)
    - [Filtering Items](#filtering-items)
      - [Adding Filters](#adding-filters)
      - [Resetting Filters](#resetting-filters)
    - [Sorting Items](#sorting-items)
    - [Adding New Rooms](#adding-new-rooms)
    - [Removing Rooms](#removing-rooms)
    - [Backing Up Database](#backing-up-database)
    - [Exporting Filtered Data](#exporting-filtered-data)
    - [Show/Hide Section](#showhide-section)
  - [Troubleshooting](#troubleshooting)
    - [Feature not working as expected](#feature-not-working-as-expected)
    - [I added a new row to the new product/filters table](#i-added-a-new-row-to-the-new-productfilters-table)


## Overview
The purpose of the Inventory Tracker is to simplify the process of managing available products and equipment. This document outlines the usage of all features of the tracker, and how to use it effectively.


## Layout
This section outlines the layout of the application, including the purpose of each section and the content that can be found within it.


### Management Buttons
The management section contains buttons to control the application, including hiding other sections, exporting data and managing rooms in the data table.

![Manageement Buttons](/images/Management-Button-Section.PNG)

### Filtering Section
The filtering section contains the filter inputs for the data table, and the sorting options for the data table. All fields within this section should automatically populate their dropdowns with all unique values found in the corresponding column in the data table.

![Filtering Section](/images/Filters-Section.PNG)

### New Item Section
The new item section contains a table used to insert new items in to the data table. Like the filtering section, all fields should automatically populate their dropdowns with all unique values from the corresponding column in the data table.

![New Item Section](/images/New-Item-Section.PNG)

### Search Section
The search section contains the value input to search the data table with, and a dropdown that controls what column is searched.

![Search Section](/images/Search-Section.PNG)

### Data Section

![Data Section](/images/Data-Section.PNG)

The data section contains all of the data stored, and is filtered based on the other sections in the application. The table contains some non-descript columns that can be used for entering specific item information, please see below for details:
- 	Description
    -	Related information, such as the quantity of the item in a pack
-	Size
    -	The physical size of the item, either as volume or physical dimensions

## Features
### Adding New Items
1.	Enter the relevant data of the new item into the blue cells of the new item table

![Enter Data](/images/New-Item.PNG)

2.	Select the “Add” button to add the data to the data table

![Click Add](/images/Add-Button.PNG)

3.	Insert the quantity into the new item in the data table

![New Quantity](/images/New-Item-Quantity.PNG)

### Editing Existing Items
1.	Locate the item in the data table either manually or utilising the search/filters sections.

![Locate Item](/images/Edit-Item.png)

2.	Make any changes required.

### Deleting Items
1.	Locate the item in the data table either manually or utilising the search/filters sections.

![Locate Item](/images/Delete-Item-Locate.PNG)

2.	Select a cell in the item row, and right click.

![Select Item](/images/Delete-Item-Right-Click.png) 

3.	Select “Delete”, then either “Delete Row” or “Delete Sheet Row”.

![Delete Item](/images/Delete-Item-Delete.png)

### Searching Items
1.	Enter a search term into the search cell.

![Search Value](/images/Search-Term.PNG)

2.	Select the column to search in the field cell.

![Search Field](/images/Search-Field.PNG)

### Filtering Items
#### Adding Filters
1.	Select the values you would like to see in the corresponding column.

![Add Filters](/images/Filters-Select.png)

#### Resetting Filters
1.	Click the “Reset Filters” button.

![Reset Filters](/images/Filters-Reset.PNG) 

### Sorting Items
1.	Select field to sort by.

![Sort Field](/images/Sort-Field.png)

2.	Select direction to sort data.

![Sort Direction](/images/Sort-Direction.png)

### Adding New Rooms
1.	Click the “New Room” button.

![New Room Button](/images/New-Room-Button.PNG)

2.	Enter the name or room code of the new room and click “OK”. E.G. G101

![New Room Name](/images/New-Room-Name.PNG)

### Removing Rooms
1.	Click the “Delete Room” button.

![Delete Room Button](/images/Delete-Room-Button.PNG)

2.	Enter the room name or code as it appears in the data table and click “OK”. E.G. G101

![Delete Rooom Name](/images/Delete-Room-Name.PNG)

3.	Click Yes

![Yes Button](/images/Delete-Room-Confirm.PNG)

### Backing Up Database
1.	Click the “Create Backup” button.

![Backup Button](/images/Backup-Button.PNG)

2.	Enter a file name and choose where to save the backup.

![Backup Filename](/images/Backup-Location.PNG)

3.	Click “Save”.

![Save Button](/images/Backup-Save.PNG)

### Exporting Filtered Data
1.	Filter data as required.

![Filter Data](/images/Export-Filter.PNG)

2.	Click the “Export Data” button.

![Export Button](/images/Export-Button.PNG)

3.	Enter a file name and choose where to save the file.

![Export Filename](/images/Export-Location.PNG)

4.	Click “Save”.

![Save Button](/images/Export-Save.PNG)

### Show/Hide Section
1.	Click the “Show/Hide” section button for the section you want to show or hide.

![Show Hide Button](/images/Show-Hide-Section-Buttons.PNG)
 
## Troubleshooting
### Feature not working as expected
It shouldn’t occur, but if an error occurs it may cause the application to stop receiving events. This will result in selected options and buttons not function as they should. To resolve this issue, save any changes you have made, close the application and re-open it. This will re-enable events and resolve the issue.

### I added a new row to the new product/filters table
In the event of a new row being added to the new product or the filters table, please follow the below directions to restore the tables.

1.	Select a cell inside of the table that has been expanded.

![Select Row](/images/Select-Extra-Row.PNG)

2.	Select the “Design” tab at the top of the page.

![Select Design Tab](/images/Design-Tab.PNG)

3.	Select “Resize Table” to the left of the ribbon.

![Resiez Table](/images/Resize-Table.PNG)

4.	After the prompt opens, select the following based on the table that has been expanded.

New Product Table – Table headers and the first row.

![Select Table Range](/images/Resize-Table-New-Item.PNG)

Filters Table – Table headers and the first two rows.

![Select Table Range](/images/Resize-Table-Filters.PNG)

5.	Click the “OK” button.

![Ok Button](/images/Resize-Table-OK.PNG)

6.	Select the extra cells beneath the table.

![Select Extra Cells](/images/Select-Extra-Cells.PNG)

7.	Set the background colour to white, and the borders to no borders.

![Cell Styles](/images/Resize-Table-Style.PNG)