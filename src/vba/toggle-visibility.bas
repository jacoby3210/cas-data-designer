Attribute VB_Name = "Module1"
Sub ToggleVisibilityColumnsWithLID()

    ' Переменные окружения
    Dim ws As Worksheet
    Dim settings As ListObject
    Dim column As ListColumn
    Dim newState As Boolean
    
    ' Доступ к настройкам документа
    Set settings = ActiveWorkbook.Sheets("@core").ListObjects("settings")
    Set column = settings.ListColumns("show_lid_columns")
    
    ' Переключение состояния видимости
    newState = Not column.DataBodyRange.Cells(1, 1).Value
    column.DataBodyRange.Cells(1, 1).Value = newState
    
    ' Перебираем все листы в книге
    For Each ws In ActiveWorkbook.Sheets
        Dim table As ListObject
        For Each table In ws.ListObjects
            For Each column In table.ListColumns
                ' Проверка, содержит ли имя столбца ":lid"
                If InStr(1, column.Name, ":lid") Then
                    column.Range.EntireColumn.Hidden = Not newState
                End If
            Next column
        Next table
    Next ws

End Sub