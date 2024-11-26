
' Toggle the display of service (‘helper’) columns
Sub ToggleVisibilityMode()

  ' Environment variables
  Dim ws As Worksheet
  Dim settings As ListObject
  Dim column As ListColumn
  Dim newState As Boolean
  
  ' Access the document settings
  Set settings = ActiveWorkbook.Sheets("@core").ListObjects("settings")
  Set column = settings.ListColumns("show_lid_columns")
  
  ' Toggle the visibility state
  newState = Not column.DataBodyRange.Cells(1, 1).Value
  column.DataBodyRange.Cells(1, 1).Value = newState
  
  ' Iterate through all sheets in the workbook
  For Each ws In ActiveWorkbook.Sheets
    Dim table As ListObject
    For Each table In ws.ListObjects
      For Each column In table.ListColumns
        ' Check if the column name contains ":lid"
        If InStr(1, column.name, ":lid") Or InStr(1, column.name, "sig") Then
          column.Range.EntireColumn.Hidden = Not newState
        End If
      Next column
    Next table
  Next ws

End Sub