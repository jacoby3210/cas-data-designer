' Get the range to search for a value
Function GetSearchRange() As Range
    ' Check if the active cell has Data Validation
    On Error Resume Next ' Handle cases where Validation is not present
    validationFormula = activeCell.Validation.Formula1
    On Error GoTo 0 ' Restore error handling
    
    ' Ensure validation is present
    If validationFormula = "" Then
        MsgBox "The active cell does not have Data Validation!", vbExclamation, "No Data Validation"
        Exit Function
    End If
    
    ' Check if validation is a range
    If Not Left(validationFormula, 1) = "=" Then
        MsgBox "Data Validation is not linked to a range.", vbExclamation, "Invalid Validation"
        Exit Function
    End If
     
    Set GetSearchRange = Range(Mid(validationFormula, 2))
    
End Function

' Get a target cell with a search value
Function GetTargetCell() As Range

    ' Setup the workflow
    Dim searchRange As Range
    Dim targetÑell As Range
    
    ' Search for the value in the range
    Set searchRange = GetSearchRange()
    Set targetÑell = searchRange.Find(What:=activeCell.Value, LookIn:=xlValues, LookAt:=xlWhole)
        
    ' Check if the value was found
    If Not targetÑell Is Nothing Then
        ' Select the found cell and show its address
        targetÑell.Select
        MsgBox "Value found at: " & targetÑell.Address, vbInformation, "Match Found"
    Else
        MsgBox "Value not found in the validation range.", vbExclamation, "No Match Found"
    End If
    
    Set GetTargetCell = targetÑell
    
End Function

' Select the row to which the target cell belongs
Function SelectTableRow(targetCell As Range)

    ' Setup the workflow
    Dim table As ListObject
    Dim row As ListRow

    ' Ensure the target cell is valid
    If targetCell Is Nothing Then
        MsgBox "Target cell is not specified!", vbExclamation, "Error"
        Exit Function
    End If

    ' Check if the target cell belongs to a table
    On Error Resume Next
    Set table = targetCell.ListObject
    On Error GoTo 0

    If table Is Nothing Then
        MsgBox "The target cell does not belong to a table!", vbExclamation, "Error"
        Exit Function
    End If

    ' Get the table row that contains the target cell
    Set row = table.ListRows(targetCell.row - table.HeaderRowRange.row)

    ' Select the entire row of the table
    If Not row Is Nothing Then
        row.Range.Select
        MsgBox "Selected row: " & row.Index, vbInformation, "Row Selected"
    Else
        MsgBox "The target cell is not within the data body of the table!", vbExclamation, "Error"
    End If
End Function

' Go to the source cell from the data validation
Sub ButtonGoToSourceCell()

    Dim targetCell As Range
    Set targetCell = GetTargetCell()
    Call SelectTableRow(targetCell)
    
End Sub

