' Get the ID for the next row
Function GetNextTableID(table As ListObject) As Integer

    ' Setup Workflow
    Dim name As String
    Dim ai As Integer
    Dim pos As Integer

    ' Extract counter from column name and increment it by one
    name = table.ListColumns(1).name
    pos = InStr(1, name, ":")
    ai = CInt(Trim(Mid(name, pos + 1))) + 1
    table.ListColumns(1).name = Left(name, pos - 1) + ":" + CStr(ai)

    ' Return the current counter value
    GetNextTableID = ai

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

' Get the locale ID for the next row
Function GetNextLocaleID() As Integer
    
    ' Setup Workflow
    Dim settings As ListObject
    Dim column As ListColumn

    ' Maintaince
    Set settings = ActiveWorkbook.Sheets("@core").ListObjects("settings")
    Set column = settings.ListColumns("ai_counter_locale_table")
    column.DataBodyRange.Cells(1, 1).Value = CInt(column.DataBodyRange.Cells(1, 1).Value) + 1
    GetNextLocaleID = column.DataBodyRange.Cells(1, 1).Value

End Function

' Get the locale ID for the next row
Function SetNextLocaleID(ai As Integer) As Integer
    
    ' Setup Workflow
    Dim settings As ListObject
    Dim column As ListColumn

    ' Maintaince
    Set settings = ActiveWorkbook.Sheets("@core").ListObjects("settings")
    Set column = settings.ListColumns("ai_counter_locale_table")
    column.DataBodyRange.Cells(1, 1).Value = ai

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
Function AddRows(table As ListObject, count As Integer)
    
    ' Setup workflow
    Dim dataArray() As Variant
    Dim defaultValue As Variant
    Dim row As Integer, column As Integer
    Dim nextTableID As Integer: nextTableID = GetNextTableID(table)
    Dim nextLocaleID As Integer: nextLocaleID = GetNextLocaleID()

    ' Resize the array to hold all the rows we need to add
    ReDim dataArray(1 To count, 1 To table.ListColumns.count)

    ' Add specified number of rows
    For row = 1 To count
        table.ListRows.Add
        For column = 1 To table.ListColumns.count
        
            defaultValue = GetDefaultValue(table, column)
            Select Case column
                Case 1
                    ' Set value for the first column (ID)
                    dataArray(row, column) = nextTableID
                    nextTableID = nextTableID + 1
                Case 2
                    dataArray(row, column) = defaultValue + CStr(nextTableID - 1)
                Case Else
                    If InStr(1, table.ListColumns(column).name, ":lid") > 0 Then
                        dataArray(row, column) = nextLocaleID
                        nextLocaleID = nextLocaleID + 1
                    Else
                        dataArray(row, column) = defaultValue
                    End If
            End Select
        Next column
    Next row
    
    table.DataBodyRange.Cells(table.ListRows.count - count + 1, 1).Resize(count, table.ListColumns.count).Value = dataArray
    SetNextTableID table, nextTableID
    SetNextLocaleID nextLocaleID

End Function

' Add multiple entities to the current table from button
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


' Add multiple entities to the current table from button
Sub ButtonSimple()
    
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