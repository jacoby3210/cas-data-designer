' Check if the rows in the collection are consecutive
Function AreSequentialRows(selectedRows As Range, offset As Long) As Boolean

    ' Setup the workflow
    Dim selectedRowIndices As Collection
        
    ' Collect row numbers of selected cells
    Set selectedRowIndices = New Collection
    For Each cell In selectedRows.Rows
        selectedRowIndices.Add cell.row - offset
    Next cell
    
    Dim i As Integer
    For i = 2 To selectedRowIndices.count
        If selectedRowIndices(i) <> selectedRowIndices(i - 1) + 1 Then
            AreSequentialRows = False
            Exit Function
        End If
    Next i
    AreSequentialRows = True

End Function

' Basic function for moving rows
Function RowMove(offset As Integer)

    ' Setup the workflow
    Dim table As ListObject
    Dim selectedRows As Range
    Dim firstSelectedRowIndex As Long
    Dim lastSelectedRowIndex As Long
    Dim targetRowIndex As Long

    On Error GoTo ErrorHandler

    ' Get the selected cells and table
    Set table = Selection.ListObject
    If table Is Nothing Then
        Err.Raise 1001, "RowMove", "The selection must be within a table!"
    End If

    ' Determine the selected rows and validate the selection
    Set selectedRows = Intersect(Selection, table.DataBodyRange)
    If selectedRows Is Nothing Then
        Err.Raise 1002, "RowMove", "The selection must include rows of the table!"
    End If

    ' Check if rows are consecutive
    If Not AreSequentialRows(selectedRows, table.HeaderRowRange.row) Then
        Err.Raise 1003, "RowMove", "Selected rows must be consecutive!"
    End If

    ' Get the first and last selected row within the table
    firstSelectedRowIndex = selectedRows.Rows(1).row - table.HeaderRowRange.row
    lastSelectedRowIndex = table.ListRows.count

    ' Check if the rows can be moved
    If firstSelectedRowIndex + offset < 1 Or firstSelectedRowIndex + offset + selectedRows.Rows.count - 1 > lastSelectedRowIndex Then
        Err.Raise 1004, "RowMove", "Cannot move rows beyond the table boundaries!"
    End If

    ' Determine the target row for the cut range
    If offset > 0 Then
        ' Moving down: Insert after the target range
        targetRowIndex = firstSelectedRowIndex + offset + selectedRows.Rows.count
    Else
        ' Moving up: Insert before the target range
        targetRowIndex = firstSelectedRowIndex + offset
    End If

    ' Move rows
    Application.ScreenUpdating = False
    table.DataBodyRange.Rows(firstSelectedRowIndex).Resize(selectedRows.Rows.count).Cut
    table.ListRows(targetRowIndex).Range.Insert Shift:=xlDown
    Application.ScreenUpdating = True

    Exit Function

ErrorHandler:
    MsgBox "Error: " & Err.Description, vbExclamation, "Error"
    Err.Clear
End Function

' Button for moving rows up
Sub ButtonMoveRowsUp()
    RowMove -1
End Sub

' Button for moving rows down
Sub ButtonMoveRowsDown()
    RowMove 1
End Sub