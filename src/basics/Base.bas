' Main procedure to create a sheet and a table
Sub ButtonCreateWorkSheet()

    ' Setup workflow
    Dim ws As Worksheet
    Dim name As String
    
    ' Get the new table name from the user
    name = InputBox("Enter the name of the new table:", "Table Name", "NewTable")
    If name = "" Then Exit Sub ' Exit if no input is provided
    
    ' Check if the sheet already exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(name)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        MsgBox "Sheet '" & name & "' already exists!", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Create a new sheet
    Set ws = CreateSheet(name)
    
    ' Create the table on the new sheet
    CreateTable ws, name, "A1"
    
    ' Notify the user
    MsgBox "Table created on sheet '" & ws.name & "'!", vbInformation, "Success"
End Sub

' Main procedure to create a table on current sheet
Sub ButtonCreateTable()
    
    ' Setup workflow
    Dim ws As Worksheet
    Dim name As String
    Dim lastRow As Long
    Dim startCell As Range
    
    ' Get the new table name from the user
    name = InputBox("Enter the name of the new table:", "Table Name", "NewTable")
    If name = "" Then Exit Sub ' Exit if no input is provided
    
    Set ws = ActiveSheet
    
    ' Find the last row with data on the sheet
    lastRow = ws.Cells(ActiveSheet.rows.count, 1).End(xlUp).row

    ' Define the starting cell for the new table (add 2 rows for spacing)
    Set startCell = ws.Cells(lastRow + 2, 1)
    
    ' Create the table on the new sheet
    CreateTable ws, name, startCell.address
       
    ' Notify the user
    MsgBox "Table created on sheet '" & ws.name & "'!", vbInformation, "Success"
    
End Sub

' Function to create a new sheet
Function CreateSheet(sheetName As String) As Worksheet

    ' Setup workflow
    Dim ws As Worksheet
    
    ' Create a new sheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.name = sheetName
    
    ' Return the newly created worksheet
    Set CreateSheet = ws
End Function

Function CreateTable(ws As Worksheet, tableName As String, cellRootAddress As String)

    ' Setup workflow
    Dim tbl As ListObject
    Dim rng As Range
    Dim data As Variant
    Dim columnWidths As Variant
    Dim row As Integer, col As Integer
    Dim startRow As Long, startCol As Long
    
    ' Parse the starting cell address
    startRow = ws.Range(cellRootAddress).row
    startCol = ws.Range(cellRootAddress).column
    
    ' Sample data to populate the table
    data = Array( _
        Array("id:1", "label:label", "name:lid", "name:ltext", "desc:lid", "desc:ltext", "note:lid", "note:ltext", "sig:formula"), _
        Array("0", "ENTITY_", "'-", "Èìÿ", "'-", "Îïèñàíèå", "'-", "Ïðèìå÷àíèå", "=CONCAT(A1;"" : "";B1)") _
    )
    
    ' Place the data into the sheet starting from the specified address
    For row = LBound(data) To UBound(data)
        For col = LBound(data(row)) To UBound(data(row))
            ws.Cells(startRow + row, startCol + col).Value = data(row)(col) ' Place each value in the corresponding cell
        Next col
    Next row
    
    ' Define the range for the table based on the data size
    Set rng = ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow + UBound(data), startCol + UBound(data(0))))
    
    ' Create the table
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.name = tableName
    
    ' Set column widths in units
    columnWidths = Array(10, 25, 10, 10, 10, 50, 10, 50) ' Specify widths for each column
    
    For col = LBound(columnWidths) To UBound(columnWidths)
        If col + 1 <= tbl.ListColumns.count Then ' Prevent errors if more widths are defined than columns
            tbl.Range.Columns(col + 1).ColumnWidth = columnWidths(col)
        End If
    Next col

End Function