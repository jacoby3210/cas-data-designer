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

' Get the ID for the next row
Function GetNextID(table As ListObject) As Integer
  Dim name As String
  Dim ai As Integer
  Dim pos As Integer

  ' Extract counter from column name and increment it by one
  name = table.ListColumns(1).Name
  pos = InStr(1, name, ":")
  ai = CInt(Trim(Mid(name, pos + 1))) + 1
  table.ListColumns(1).Name = Left(name, pos - 1) + ":" + CStr(ai)

  ' Return the current counter value
  GetNextID = ai
End Function

' Get the locale ID for the next row
Function GetNextLocaleID() As Integer
  Dim settings As ListObject
  Dim column As ListColumn

  Set settings = ActiveWorkbook.Sheets("@core").ListObjects("settings")
  Set column = settings.ListColumns("ai_counter_locale_table")
  column.DataBodyRange.Cells(1, 1).Value = CInt(column.DataBodyRange.Cells(1, 1).Value) + 1
  GetNextLocaleID = column.DataBodyRange.Cells(1, 1).Value
End Function

' Add a new entry to the table
Sub AddNewEntity()
  Dim table As ListObject
  Dim column As ListColumn
  Dim targetRow As ListRow
  Dim defaultValue As Variant
  Dim cell As Range
  Dim i As Integer

  ' Access the current table
  Set table = GetTableAtCursor()
  
  ' Check if the table is selected
  If table Is Nothing Then
    MsgBox "Please select the table where you want to add a new entry.", vbInformation
    Exit Sub
  End If

  ' Add a new row
  Set targetRow = table.ListRows.Add

  i = 1
  For Each column In table.ListColumns
    Set cell = targetRow.Range(1, i)

    ' Assign values to cells in the new row
    Select Case i
      Case 1
        cell.Value = GetNextID(table)
      Case 2
        defaultValue = column.DataBodyRange.Cells(1, 1).Value
        cell.Value = defaultValue + CStr(targetRow.Range(1, 1).Value)
      Case Else
        If InStr(1, column.Name, ":lid") > 0 Then
          cell.Value = GetNextLocaleID()
        Else
          ' Copy the value from the first cell of the column
          column.DataBodyRange.Cells(1, 1).Copy cell
        End If
    End Select

      i = i + 1
  Next column

End Sub