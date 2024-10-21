Attribute VB_Name = "Module2"
' �������� ������� � ������� ��������� ������
Function GetTableAtCursor() As ListObject
    ' ���������� ���������
    Dim activeTable As ListObject
    
    ' ���������, ��������� �� �������� ������ ������ �������
    On Error Resume Next ' ���������� ������, ���� ������ �� ������ �������
    Set activeTable = ActiveCell.ListObject
    On Error GoTo 0 ' ��������������� ��������� ������
        
    ' ���������� ��������� ������� ��� Nothing, ���� ������ �� � �������
    Set GetTableAtCursor = activeTable
End Function

' �������� id ��� ��������� ������
Function GetNextID(table) As Integer
    ' ���������� ���������
    Dim name As String
    Dim ai As Integer
    Dim pos As Integer
    
    ' �������� ������� �� ����� ������� � ���������� ��� �� �������
    name = table.ListColumns(1).name
    pos = InStr(1, name, ":")
    ai = CInt(Trim(Mid(name, pos + 1))) + 1
    table.ListColumns(1).name = Left(name, pos - 1) + ":" + CStr(ai)
    
    ' ���������� ������� �������� ��������
    GetNextID = ai
    
End Function

' �������� locale id ��� ��������� ������
Function GetNextLocaleID() As Integer
    Set settings = ThisWorkbook.Sheets("@core").ListObjects("settings")
    Set column = settings.ListColumns("ai_counter_locale_table")
    column.DataBodyRange.Cells(1, 1).Value = CInt(column.DataBodyRange.Cells(1, 1).Value) + 1
    GetNextLocaleID = column.DataBodyRange.Cells(1, 1).Value
End Function

' ���������� ����� ������ � �������
Sub AddNewEntity()
    ' ���������� ��� ������ �������
    Dim table As ListObject
    Dim column As ListColumn
    Dim targetRow As ListRow
    Dim defaultValue As Variant
    Dim cell As Range
    Dim i As Integer
    
    ' ������ � ������� �������.
    Set table = GetTableAtCursor()
    Set targetRow = table.ListRows.Add
    
    i = 1
    For Each column In table.ListColumns
        Set cell = targetRow.Range(1, i)
        
        ' ���������� �������� ������� � ����� ������
        Select Case i
            Case 1
                cell.Value = GetNextID(table)
            Case 2
                defaultValue = column.DataBodyRange.Cells(1, 1).Value
                cell.Value = defaultValue + CStr(targetRow.Range(1, 1).Value)
            Case Else
                If InStr(1, column.name, ":lid") > 0 Then
                    cell.Value = GetNextLocaleID()
                Else
                    ' ����������� �������� �� ������ ������ �������
                    column.DataBodyRange.Cells(1, 1).Copy cell
                End If
        End Select
        
        i = i + 1
    Next column
    
End Sub

