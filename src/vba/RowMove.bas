' Check if the rows in the collection are consecutive
Function AreRowsConsecutive(selectedRows As Range, offset As Long) As Boolean

    ' Setup the workflow
    Dim selectedRowNums As Collection
        
    ' Collect row numbers of selected cells
    Set selectedRowNums = New Collection
    For Each cell In selectedRows.Rows
        rowNum = cell.row - offset
        selectedRowNums.Add rowNum
    Next cell
    
    Dim i As Integer
    For i = 2 To selectedRowNums.count
        If selectedRowNums(i) <> selectedRowNums(i - 1) + 1 Then
            AreRowsConsecutive = False
            Exit Function
        End If
    Next i
    AreRowsConsecutive = True

End Function

' Basic function for moving rows
Function RowMove(offset As Integer)

    ' Setup the workflow
    Dim rng As Range
    Dim tbl As ListObject
    Dim firstRow As Long
    Dim lastRow As Long
    Dim selectedRows As Range
    Dim targetRow As Long

    On Error Resume Next

    ' Get the selected cells and table
    Set rng = Selection
    Set tbl = rng.ListObject

    If tbl Is Nothing Then
        MsgBox "The selection must be within a table!", vbExclamation, "Error"
        Exit Function
    End If
    On Error GoTo 0

    ' Determine the selected rows and validate the selection
    Set selectedRows = Intersect(rng, tbl.DataBodyRange)
    If selectedRows Is Nothing Then
        MsgBox "The selection must include rows of the table!", vbExclamation, "Error"
        Exit Function
    End If

    ' Check if rows are consecutive
    If Not AreRowsConsecutive(selectedRows, tbl.HeaderRowRange.row) Then
        MsgBox "Selected rows must be consecutive!", vbExclamation, "Error"
        Exit Function
    End If

    ' Get the first and last selected row within the table
    firstRow = selectedRows.Rows(1).row - tbl.HeaderRowRange.row
    lastRow = tbl.ListRows.count

    ' Check if the rows can be moved
    If firstRow + offset < 1 Or firstRow + offset + selectedRows.Rows.count - 1 > lastRow Then
        MsgBox "Cannot move rows beyond the table boundaries!", vbExclamation, "Error"
        Exit Function
    End If

    ' Determine the target row for the cut range
    If offset > 0 Then
        ' Moving down: Insert after the target range
        targetRow = firstRow + offset + selectedRows.Rows.count
    Else
        ' Moving up: Insert before the target range
        targetRow = firstRow + offset
    End If

    ' Move rows
    Application.ScreenUpdating = False
    tbl.ListRows(firstRow).Range.Resize(selectedRows.Rows.count).Cut
    tbl.ListRows(targetRow).Range.Insert Shift:=xlDown
    Application.ScreenUpdating = True

End Function

' Button for moving rows up
Sub ButtonMoveRowsUp()
    RowMove -1
End Sub

' Button for moving rows down
Sub ButtonMoveRowsDown()
    RowMove 1
End Sub