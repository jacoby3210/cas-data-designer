Sub CreateSheetWithTable()

    ' Setup workflow
    Dim ws As Worksheet
    Dim tbl As ListObject
    Dim rng As Range
    Dim newSheetName As String
    Dim newTableName As String
    Dim data As Variant
    
    ' Get the new table name from the user
    newTableName = InputBox("Enter the name of the new table:", "Table Name", "NewTable")
    If newTableName = "" Then Exit Sub ' Exit if no input is provided
    
    newSheetName = newTableName
    
    ' Check if the sheet already exists
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(newSheetName)
    On Error GoTo 0
    
    If Not ws Is Nothing Then
        MsgBox "Sheet '" & newSheetName & "' already exists!", vbExclamation, "Error"
        Exit Sub
    End If
    
    ' Create a new sheet
    Set ws = ThisWorkbook.Sheets.Add
    ws.name = newSheetName
    
    ' Sample data to populate the table
    data = Array( _
        Array("id:1", "label", "name:lid", "name:ltext", "desc:lid", "desc:ltext", "note:lid", "note:ltext", "sig:formula"), _
        Array("0", "ENTITY_", "'-", "Èìÿ", "'-", "Îïèñàíèå", "'-", "Ïðèìå÷àíèå", "=CONCAT(A1;"" : "";B1)") _
    )
    
    ' Place the data into the sheet
    For row = LBound(data) To UBound(data)
        For col = LBound(data(row)) To UBound(data(row))
            ws.Cells(row + 1, col + 1).Value = data(row)(col) ' Place each value in the corresponding cell
        Next col
    Next row
    
    
    ' Define the range for the table
    Set rng = ws.Range("A1").CurrentRegion
    
    ' Create the table
    Set tbl = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    tbl.name = newTableName
    
    ' Set column widths in units
    Dim columnWidths As Variant
    columnWidths = Array(10, 25, 10, 10, 10, 50, 10, 50) ' Specify widths for each column
    
    For col = LBound(columnWidths) To UBound(columnWidths)
        tbl.Range.Columns(col + 1).ColumnWidth = columnWidths(col)
    Next col
    
    ' Notify the user
    MsgBox "Table created on sheet '" & ws.name & "'!", vbInformation, "Success"

End Sub


