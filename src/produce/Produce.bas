Sub RunEmbeddedPythonScript()

    ' Setup Workflow
    Dim fso As Object
    Dim tempScriptPath As String
    Dim shellCommand As String

    ' Path to temporary file
    workbookPath = ThisWorkbook.Path
    If workbookPath = "" Then
        MsgBox "Êíèãà íå ñîõðàíåíà. Ñîõðàíèòå êíèãó ïåðåä çàïóñêîì ñêðèïòà.", vbExclamation
        Exit Sub
    End If
    Call SaveIfUnsaved
    
    ' Create the file and write the Python code
    tempScriptPath = workbookPath & "\temp_script.py"
    Set fso = CreateObject("Scripting.FileSystemObject")
    With fso.CreateTextFile(tempScriptPath, True)
        .Write GetPythonScript()
        .Close
    End With

    ' Run the Python script
    shellCommand = "python """ & tempScriptPath & """"
    Shell shellCommand, vbNormalFocus

    ' Optional: Delete the temporary file
    fso.DeleteFile tempScriptPath
    
End Sub

Function GetPythonScript() As String
    
    ' Setup Workflow
    Dim pythonCode As String
    
    ' Initialize Python code as empty
    pythonCode = ""

    ' Append Python code in smaller parts
    pythonCode = pythonCode & "import json" & vbCrLf
    pythonCode = pythonCode & "import openpyxl" & vbCrLf
    pythonCode = pythonCode & "import os" & vbCrLf & vbCrLf
    
    'Handle table header
    pythonCode = pythonCode & "def get_column_indices(headers):" & vbCrLf
    pythonCode = pythonCode & "    column_indices = {" & vbCrLf
    pythonCode = pythonCode & "        'data': []," & vbCrLf
    pythonCode = pythonCode & "        'locales': []," & vbCrLf
    pythonCode = pythonCode & "        'refs': []" & vbCrLf
    pythonCode = pythonCode & "    }" & vbCrLf
    pythonCode = pythonCode & "    column_names = []" & vbCrLf & vbCrLf
    pythonCode = pythonCode & "    for i, header in enumerate(headers):" & vbCrLf
    pythonCode = pythonCode & "        column_name, column_type = header.split(':')[0], header.split(':')[1].lower()" & vbCrLf
    pythonCode = pythonCode & "        column_names.append(column_name)" & vbCrLf & vbCrLf
    pythonCode = pythonCode & "        if 'formula' in column_type: continue" & vbCrLf
    pythonCode = pythonCode & "        elif 'ltext' in column_type: column_indices['locales'].append(i)" & vbCrLf
    pythonCode = pythonCode & "        elif 'ref' in column_type: column_indices['refs'].append(i)" & vbCrLf
    pythonCode = pythonCode & "        else: column_indices['data'].append(i)" & vbCrLf & vbCrLf
    pythonCode = pythonCode & "    return column_names, column_indices" & vbCrLf & vbCrLf
    
    ' Convert table to JSON file
    pythonCode = pythonCode & "def convert_table_to_json(table_name, table_data, json_locale_data):" & vbCrLf
    pythonCode = pythonCode & "    headers = table_data[0]" & vbCrLf
    pythonCode = pythonCode & "    rows = table_data[1:]" & vbCrLf
    pythonCode = pythonCode & "    column_names, column_indices = get_column_indices(headers)" & vbCrLf & vbCrLf
    pythonCode = pythonCode & "    json_table_data = {}" & vbCrLf
    pythonCode = pythonCode & "    for row in rows:" & vbCrLf
    pythonCode = pythonCode & "        row_data = {column_names[i]: row[i] for i in column_indices['data'] if row[i] is not None}" & vbCrLf
    pythonCode = pythonCode & "        ref_data = {column_names[i]: int(row[i].split(':')[0]) for i in column_indices['refs'] if row[i] is not None}" & vbCrLf
    pythonCode = pythonCode & "        json_table_data[row[0]] = {**row_data, **ref_data}" & vbCrLf & vbCrLf
    pythonCode = pythonCode & "        for i in column_indices['locales']:" & vbCrLf
    pythonCode = pythonCode & "            json_locale_data[row[i-1]] = {'text': row[i], 'column': column_names[i], 'table': table_name}" & vbCrLf & vbCrLf
    pythonCode = pythonCode & "    json_path = os.path.join(output_dir, f'{table_name}.json')" & vbCrLf
    pythonCode = pythonCode & "    with open(json_path, 'w', encoding='utf-8') as json_file:" & vbCrLf
    pythonCode = pythonCode & "        json.dump(json_table_data, json_file, ensure_ascii=False, indent=2)" & vbCrLf
    pythonCode = pythonCode & "    print(f'Table {table_name} successfully transformed into JSON')" & vbCrLf & vbCrLf

    ' Convert workbook to JSON folder
    pythonCode = pythonCode & "def parse_wb(wb_path):" & vbCrLf
    pythonCode = pythonCode & "    json_locale_data = {}" & vbCrLf
    pythonCode = pythonCode & "    json_locale_path = os.path.join(output_dir, 'locale.json')" & vbCrLf
    pythonCode = pythonCode & "    if os.path.exists(json_locale_path):" & vbCrLf
    pythonCode = pythonCode & "        with open(json_locale_path, 'r', encoding='utf-8') as file:" & vbCrLf
    pythonCode = pythonCode & "            json_locale_data = json.load(file)" & vbCrLf
    pythonCode = pythonCode & "    wb_handler = openpyxl.load_workbook(wb_path, keep_links=True, keep_vba=True)" & vbCrLf
    pythonCode = pythonCode & "    for wb_sheet_name in wb_handler.sheetnames:" & vbCrLf
    pythonCode = pythonCode & "        if wb_sheet_name.startswith('@'):" & vbCrLf
    pythonCode = pythonCode & "            print(f'Sheet {wb_sheet_name} skipped (starts with the @ symbol)')" & vbCrLf
    pythonCode = pythonCode & "            continue" & vbCrLf
    pythonCode = pythonCode & "        wb_sheet_handler = wb_handler[wb_sheet_name]" & vbCrLf
    pythonCode = pythonCode & "        for table_name, table in wb_sheet_handler.tables.items():" & vbCrLf
    pythonCode = pythonCode & "            table_range = wb_sheet_handler[table]" & vbCrLf
    pythonCode = pythonCode & "            table_data = [[cell.value for cell in row] for row in table_range]" & vbCrLf
    pythonCode = pythonCode & "            # Assuming convert_table_to_json is defined elsewhere" & vbCrLf
    pythonCode = pythonCode & "            convert_table_to_json(table_name, table_data, json_locale_data)" & vbCrLf
    pythonCode = pythonCode & "    with open(json_locale_path, 'w', encoding='utf-8') as json_file:" & vbCrLf
    pythonCode = pythonCode & "        json.dump(json_locale_data, json_file, ensure_ascii=False, indent=2)" & vbCrLf & vbCrLf


    ' Compile tables from the book
    pythonCode = pythonCode & "input_dir = os.path.dirname(r'" & ThisWorkbook.FullName & "')" & vbCrLf
    pythonCode = pythonCode & "output_dir = os.path.dirname(r'" & ThisWorkbook.FullName & "')" & vbCrLf
    pythonCode = pythonCode & "wb_file_path = r'" & ThisWorkbook.FullName & "'" & vbCrLf
    pythonCode = pythonCode & "if wb_file_path.endswith('.xlsm') or wb_file_path.endswith('.xls'):" & vbCrLf
    pythonCode = pythonCode & "    print(f'File detected {wb_file_path}')" & vbCrLf
    pythonCode = pythonCode & "    parse_wb(wb_file_path)" & vbCrLf & vbCrLf
    
    pythonCode = pythonCode & "print ('All files have been processed.')" & vbCrLf
    pythonCode = pythonCode & "input('Press Enter to exit...')" & vbCrLf & vbCrLf
    pythonCode = pythonCode & "    "
    
    ' Append more Python code here...

    GetPythonScript = pythonCode
    
End Function

Function SaveIfUnsaved()
    ' Check if there are unsaved changes
    If Not ThisWorkbook.Saved Then
        ' Prompt the user to save the workbook
        Dim response As VbMsgBoxResult
        response = MsgBox("There are unsaved changes. Do you want to save the workbook?", vbYesNoCancel + vbQuestion, "Save Changes")

        Select Case response
            Case vbYes
                ' Save the workbook
                ThisWorkbook.Save
                MsgBox "Workbook saved successfully!", vbInformation
            Case vbNo
                ' Do nothing
                MsgBox "Changes not saved.", vbExclamation
            Case vbCancel
                ' Cancel the operation
                MsgBox "Operation canceled.", vbExclamation
                Exit Function
        End Select
    Else
        MsgBox "No changes to save.", vbInformation
    End If
End Function