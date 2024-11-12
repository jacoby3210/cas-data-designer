Attribute VB_Name = "Module2"

' Получить таблицу, в которой находится курсор
Function GetTableAtCursor() As ListObject
    Dim activeTable As ListObject

    ' Проверяем, находится ли активная ячейка внутри таблицы
    On Error Resume Next ' Игнорируем ошибку, если ячейка не внутри таблицы
    Set activeTable = ActiveCell.ListObject
    On Error GoTo 0 ' Восстанавливаем обработку ошибок

    ' Возвращаем найденную таблицу или Nothing, если курсор не в таблице
    Set GetTableAtCursor = activeTable
End Function

' Получить id для следующей строки
Function GetNextID(table As ListObject) As Integer
    Dim name As String
    Dim ai As Integer
    Dim pos As Integer

    ' Извлекаем счётчик из имени столбца и увеличиваем его на единицу
    name = table.ListColumns(1).Name
    pos = InStr(1, name, ":")
    ai = CInt(Trim(Mid(name, pos + 1))) + 1
    table.ListColumns(1).Name = Left(name, pos - 1) + ":" + CStr(ai)

    ' Возвращаем текущее значение счётчика
    GetNextID = ai
End Function

' Получить locale id для следующей строки
Function GetNextLocaleID() As Integer
    Dim settings As ListObject
    Dim column As ListColumn

    Set settings = ActiveWorkbook.Sheets("@core").ListObjects("settings")
    Set column = settings.ListColumns("ai_counter_locale_table")
    column.DataBodyRange.Cells(1, 1).Value = CInt(column.DataBodyRange.Cells(1, 1).Value) + 1
    GetNextLocaleID = column.DataBodyRange.Cells(1, 1).Value
End Function

' Добавление новой записи в таблицу
Sub AddNewEntity()
    Dim table As ListObject
    Dim column As ListColumn
    Dim targetRow As ListRow
    Dim defaultValue As Variant
    Dim cell As Range
    Dim i As Integer

    ' Доступ к текущей таблице
    Set table = GetTableAtCursor()
    
    ' Проверка, выбрана ли таблица
    If table Is Nothing Then
        MsgBox "Пожалуйста, выберите таблицу, в которой хотите добавить новую запись.", vbInformation
        Exit Sub
    End If

    ' Добавление новой строки
    Set targetRow = table.ListRows.Add

    i = 1
    For Each column In table.ListColumns
        Set cell = targetRow.Range(1, i)

        ' Присвоение значений ячейкам в новой строке
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
                    ' Копирование значения из первой ячейки столбца
                    column.DataBodyRange.Cells(1, 1).Copy cell
                End If
        End Select

        i = i + 1
    Next column

End Sub