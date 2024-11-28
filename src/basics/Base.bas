' Main procedure: create a sheet and a table
Sub ButtonCreateWorkSheet()

    On Error GoTo ErrorHandler ' Set up error handling

    ' Workflow setup
    Dim ws As Worksheet
    Dim name As String
    
    ' Get the new table name from the user
    name = InputBox("Enter the name of the new table:", "Table Name", "NewTable")
    If name = "" Then Exit Sub ' Exit if no input is provided
    
    ' Check if the sheet already exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(name)
    On Error GoTo ErrorHandler
    
    If Not ws Is Nothing Then
        Err.Raise vbObjectError + 1, "ButtonCreateWorkSheet", _
            "Sheet '" & name & "' already exists!"
    End If
    
    ' Create a new sheet
    Set ws = CreateSheet(name)
    
    ' Create the table on the new sheet
    CreateTable ws, name, "A1"
    
    ' Notify the user
    ' MsgBox "Table created on sheet '" & ws.name & "'!", vbInformation, "Success"
    Exit Sub ' Exit the procedure normally

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error in ButtonCreateWorkSheet"
    Err.Clear ' Clear the error
End Sub

' Main procedure: create a table on the current sheet
Sub ButtonCreateTable()

    On Error GoTo ErrorHandler ' Set up error handling

    ' Workflow setup
    Dim ws As Worksheet
    Dim name As String
    Dim lastRow As Long
    Dim startCell As Range
    
    ' Get the new table name from the user
    name = InputBox("Enter the name of the new table:", "Table Name", "NewTable")
    If name = "" Then Exit Sub ' Exit if no input is provided
    
    ' Use the active sheet
    Set ws = ActiveSheet
    
    ' Find the last row with data
    lastRow = ws.Cells(ws.rows.count, 1).End(xlUp).row
    Set startCell = ws.Cells(lastRow + 2, 1)
    startCell.Select ' Move cursor to the starting cell
    
    ' Create the table
    CreateTable ws, name, startCell.Address
    
    ' Notify the user
    ' MsgBox "Table created on sheet '" & ws.name & "'!", vbInformation, "Success"
    Exit Sub ' Exit the procedure normally

ErrorHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error in ButtonCreateTable"
    Err.Clear ' Clear the error
End Sub

' Function to create a new sheet
Function CreateSheet(sheetName As String) As Worksheet

    Dim ws As Worksheet

    ' Create a new sheet
    Set ws = ThisWorkbook.Sheets.Add
    On Error Resume Next
    ws.name = sheetName
    If Err.Number <> 0 Then
        Err.Raise vbObjectError + 2, "CreateSheet", _
                  "Failed to create sheet with name '" & sheetName & "'."
    End If
    On Error GoTo 0
    
    ' Return the newly created worksheet
    Set CreateSheet = ws
End Function

' Function to create a new table on a given Worksheet
Function CreateTable(ws As Worksheet, tableName As String, cellRootAddress As String)

    Dim tbl As ListObject
    Dim rng As Range
    Dim data As Variant
    Dim columnWidths As Variant
    Dim row As Integer, col As Integer
    Dim startRow As Long, startCol As Long

    ' Check if a table with the same name exists
    Set tbl = GetTableByName(ws, tableName)
    If Not tbl Is Nothing Then
        Err.Raise vbObjectError + 3, "CreateTable", _
            "Table name '" & tableName & "' already exists!"
    End If

    ' Parse the starting cell address
    startRow = ws.Range(cellRootAddress).row
    startCol = ws.Range(cellRootAddress).column

    ' Sample data to populate the table
    data = Array( _
        Array("id:1", "label:label", "name:lid", "name:ltext", "desc:lid", "desc:ltext", "note:lid", "note:ltext", "sig:formula"), _
        Array("0", "ENTITY_", "'-", "Name", "'-", "Description", "'-", "Note", "=CONCAT(A" & (startRow + 1) & ";"" : "";B" & (startRow + 1) & ")") _
    )
    
    ' Populate the data into the sheet
    For row = LBound(data) To UBound(data)
        For col = LBound(data(row)) To UBound(data(row))
            ws.Cells(startRow + row, startCol + col).Value = data(row)(col)
        Next col
    Next row
    
    ' Define the range for the table
    Set rng = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + UBound(data), startCol + UBound(data(0))))
    
    ' Create the table
    On Error Resume Next
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    If tbl Is Nothing Then
        Err.Raise vbObjectError + 4, "CreateTable", _
                  "Failed to create table. Check the range and data."
    End If
    On Error GoTo 0
    
    tbl.name = tableName
    
    ' Set column widths
    columnWidths = Array(10, 25, 10, 10, 10, 50, 10, 50)
    For col = LBound(columnWidths) To UBound(columnWidths)
        If col + 1 <= tbl.ListColumns.count Then
            tbl.Range.Columns(col + 1).ColumnWidth = columnWidths(col)
        End If
    Next col

End Function

' Function to get a table by name
Function GetTableByName(sheet As Worksheet, tableName As String) As ListObject
    Dim tbl As ListObject
    
    ' Iterate through all tables on the sheet
    For Each tbl In sheet.ListObjects
        If tbl.name = tableName Then
            Set GetTableByName = tbl ' Return the found table
            Exit Function
        End If
    Next tbl
    
    ' Return Nothing if not found
    Set GetTableByName = Nothing
End Function