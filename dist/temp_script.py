import json
import openpyxl
import os

def get_column_indices(headers):
    column_indices = {
        'data': [],
        'locales': [],
        'refs': []
    }
    column_names = []

    for i, header in enumerate(headers):
        column_name, column_type = header.split(':')[0], header.split(':')[1].lower()
        column_names.append(column_name)

        if 'formula' in column_type: continue
        elif 'ltext' in column_type: column_indices['locales'].append(i)
        elif 'ref' in column_type: column_indices['refs'].append(i)
        else: column_indices['data'].append(i)

    return column_names, column_indices

def convert_table_to_json(table_name, table_data, json_locale_data):
    headers = table_data[0]
    rows = table_data[1:]
    column_names, column_indices = get_column_indices(headers)

    json_table_data = {}
    for row in rows:
        row_data = {column_names[i]: row[i] for i in column_indices['data'] if row[i] is not None}
        ref_data = {column_names[i]: int(row[i].split(':')[0]) for i in column_indices['refs'] if row[i] is not None}
        json_table_data[row[0]] = {**row_data, **ref_data}

        for i in column_indices['locales']:
            json_locale_data[row[i-1]] = {'text': row[i], 'column': column_names[i], 'table': table_name}

    json_path = os.path.join(output_dir, f'{table_name}.json')
    with open(json_path, 'w', encoding='utf-8') as json_file:
        json.dump(json_table_data, json_file, ensure_ascii=False, indent=2)
    print(f'Table {table_name} successfully transformed into JSON')

def parse_wb(wb_path):
    json_locale_data = {}
    json_locale_path = os.path.join(output_dir, 'locale.json')
    if os.path.exists(json_locale_path):
        with open(json_locale_path, 'r', encoding='utf-8') as file:
            json_locale_data = json.load(file)
    wb_handler = openpyxl.load_workbook(wb_path, keep_links=True, keep_vba=True)
    for wb_sheet_name in wb_handler.sheetnames:
        if wb_sheet_name.startswith('@'):
            print(f'Sheet {wb_sheet_name} skipped (starts with the @ symbol)')
            continue
        wb_sheet_handler = wb_handler[wb_sheet_name]
        for table_name, table in wb_sheet_handler.tables.items():
            table_range = wb_sheet_handler[table]
            table_data = [[cell.value for cell in row] for row in table_range]
            # Assuming convert_table_to_json is defined elsewhere
            convert_table_to_json(table_name, table_data, json_locale_data)
    with open(json_locale_path, 'w', encoding='utf-8') as json_file:
        json.dump(json_locale_data, json_file, ensure_ascii=False, indent=2)

input_dir = os.path.dirname(r'W:\godot\cassiopeia\godot-cas-game-data-designer\dist\DesignDocumentProto.xlsm')
output_dir = os.path.dirname(r'W:\godot\cassiopeia\godot-cas-game-data-designer\dist\DesignDocumentProto.xlsm')
wb_file_path = r'W:\godot\cassiopeia\godot-cas-game-data-designer\dist\DesignDocumentProto.xlsm'
if wb_file_path.endswith('.xlsm') or wb_file_path.endswith('.xls'):
    print(f'File detected {wb_file_path}')
    parse_wb(wb_file_path)

print ('All files have been processed.')
input('Press Enter to exit...')

    