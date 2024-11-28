' Main procedure: add multiple entities into end of the current table
Sub ButtonMultiple()
    
    ' Setup workflow
    On Error GoTo ErrorHandler ' Enable error handling
    
    ' Access the current table
    Dim table As ListObject: Set table = GetTableAtCursor()
    
    ' Prompt the user to input the number of rows to add & validate the input
    Dim rowCountInput As Variant:
        rowCountInput = InputBox("Enter the count of rows to add:", "Add New Rows", 5)
        
    If Not IsNumeric(rowCountInput) Then
        MsgBox "Invalid input. Please enter a numeric value.", vbCritical
        Exit Sub
    End If

    ' Type assignment & Ensure rowCount is at least 1
    Dim rowCountInteger As Integer: rowCountInteger = CInt(rowCountInput)
    If rowCountInteger < 1 Then
        MsgBox "The number of rows must be at least 1.", vbInformation
        Exit Sub
    End If
    
    ' Add rows
    AddRows table, rowCountInteger
    
    MsgBox rowCountInteger & " rows added successfully.", vbInformation
    Exit Sub
    
    ' Handle errors raised in the function
ErrorHandler:
    
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub

' Main procedure: add single entite into the end of current table
Sub ButtonSingle()
    
    ' Setup workflow
    On Error GoTo ErrorHandler ' Enable error handling
    
    ' Access the current table
    Dim table As ListObject: Set table = GetTableAtCursor()
    
    
    ' Add rows
    AddRows table, 1
    
    MsgBox rowCountInteger & " rows added successfully.", vbInformation
    Exit Sub
    
    ' Handle errors raised in the function
ErrorHandler:
    
    MsgBox "Error: " & Err.Description, vbExclamation
End Sub
' Get the ID for the next row
Function GetNextTableID(table As ListObject) As Integer

    Dim name As String: name = table.ListColumns(1).name
    Dim pos As Integer: pos = InStr(1, name, ":")
    GetNextTableID = CInt(Trim(Mid(name, pos + 1)))

End Function

' Set the ID for the next row
Function SetNextTableID(table As ListObject, ai As Integer) As Integer

    ' Setup Workflow
    Dim name As String
    Dim pos As Integer

    ' Extract counter from column name and increment it by one
    name = table.ListColumns(1).name
    pos = InStr(1, name, ":")
    table.ListColumns(1).name = Left(name, pos - 1) + ":" + CStr(ai)

End Function

' Get Default value for cell in added row
Function GetDefaultValue(table As ListObject, column As Integer) As Variant
    
    ' Setup workflow
    Dim cell As Range
    Set cell = table.ListColumns(column).DataBodyRange.Cells(1, 1)
    
    ' Check for formula or value
    If cell.HasFormula Then
        If cell.FormulaR1C1 <> "" Then
            GetDefaultValue = cell.FormulaR1C1
        Else
            GetDefaultValue = cell.Formula
        End If
    Else
        GetDefaultValue = cell.Value
    End If
    
End Function

' Add multiple entities to the current table
Function AddRows(tableData As ListObject, count As Integer)
    
    ' Setup workflow
    Dim dataArray() As Variant
    Dim defaultValue As Variant
    Dim row As Integer, column As Integer
    
    Dim tableLocale As ListObject:
        Set tableLocale = ActiveWorkbook.Sheets("@core").ListObjects("locale")
        
    Dim nextTableID As Integer:
        nextTableID = GetNextTableID(tableData)
    Dim nextLocaleID As Integer:
        nextLocaleID = GetNextTableID(tableLocale)

    ' Resize the array to hold all the rows we need to add
    ReDim dataArray(1 To count, 1 To tableData.ListColumns.count)

    ' Add specified number of rows
    For row = 1 To count
        tableData.ListRows.Add
        For column = 1 To tableData.ListColumns.count
        
            defaultValue = GetDefaultValue(tableData, column)
            Select Case column
                Case 1
                    ' Set value for the first column (ID)
                    dataArray(row, column) = nextTableID
                    nextTableID = nextTableID + 1
                Case 2
                    dataArray(row, column) = defaultValue + CStr(nextTableID - 1)
                Case Else
                    If InStr(1, tableData.ListColumns(column).name, ":lid") > 0 Then
                        dataArray(row, column) = nextLocaleID
                        nextLocaleID = nextLocaleID + 1
                    Else
                        dataArray(row, column) = defaultValue
                    End If
            End Select
        Next column
    Next row
    
    tableData.DataBodyRange.Cells(tableData.ListRows.count - count + 1, 1).Resize(count, tableData.ListColumns.count).Value = dataArray
    SetNextTableID tableData, nextTableID
    SetNextTableID tableLocale, nextLocaleID
    
End Function

' Get the table where the cursor is located
Function GetTableAtCursor() As ListObject
  Dim activeTable As ListObject

  ' Check if the active cell is within a table
  On Error Resume Next ' Ignore error if cell is not inside a table
    Set activeTable = ActiveCell.ListObject
  On Error GoTo 0 ' Restore error handling

  ' Return the found table or Nothing if the cursor is not in a table
  Set GetTableAtCursor = activeTable
End Function