Attribute VB_Name = "Module1"
Sub ToggleVisibilityСolumnsWithLID()

    ' Переменные окружения
    Dim ws As Worksheet
    Dim settings As ListObject
    Dim column As ListColumn
    Dim newState As Boolean
    
    ' Доступ к настройкам документа
    Set settings = ThisWorkbook.Sheets("@core").ListObjects("settings")
    Set column = settings.ListColumns("show_lid_columns")
    newState = Not column.DataBodyRange.Cells(1, 1).Value
    column.DataBodyRange.Cells(1, 1).Value = newState
    
    ' Перебираем все листы в книге
    For Each ws In ThisWorkbook.Sheets
        For Each table In ws.ListObjects
           For Each column In table.ListColumns
               If InStr(1, column.name, ":lid") Then
                   column.Range.EntireColumn.Hidden = Not newState
               End If
           Next column
       Next table
    Next ws

End Sub


