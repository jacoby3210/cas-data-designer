' Get the table where the cursor is located
Public Function GetTableAtCursor() As ListObject

    Dim table As ListObject
    On Error Resume Next                ' Ignore error if ActiveCell is not within a table
    Set table = ActiveCell.ListObject
    On Error GoTo 0                     ' Restore error handling

    If table Is Nothing Then
        Err.Raise vbObjectError + 1000, "GetTableAtCursor", "Please select the table where you want to add new entries."
    End If

    Set GetTableAtCursor = table

End Function
