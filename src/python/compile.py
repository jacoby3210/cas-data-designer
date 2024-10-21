import csv
import json
import openpyxl
import os

# Функция для конвертации таблицы в CSV и JSON
def convert_table_to_csv_json(table_name, table_data, csv_locale_writer):

  # Подготовка заголовков (удаление текста после двоеточия)
  columns_all = table_data[0]
  column_names = []
  column_indices_with_data = []
  column_indices_with_locales = []
  column_indices_with_refs = []

  for i, header in enumerate(columns_all):
    if isinstance(header, str):
      column_names.append(header.split(":")[0])
      column_type = header.split(":")[1]
      
      if 'formula' in column_type.lower():
        continue  # Пропускаем столбцы с "formula" в заголовке
      elif 'ltext' in column_type.lower():
        column_indices_with_locales.append(i)
      elif 'ref' in column_type.lower():
        column_indices_with_refs.append(i)
      else:
        column_indices_with_data.append(i)

  # Запись данных
  json_data = []
  json_path = os.path.join(output_dir, f'{table_name}.json')
  for row in table_data[1:]:
    json_data.append(
      {column_names[i]: row[i] for i in column_indices_with_data} | 
      {column_names[i]: int(row[i].split(":")[0]) for i in column_indices_with_refs}
    )

    # Запись локализованного текста в отдельный CSV
    if csv_locale_writer and column_indices_with_locales:
      for i in column_indices_with_locales:
        csv_locale_writer.writerow([row[i-1], row[i], columns_all[i].split(":")[0], table_name])

  # Запись данных в JSON
  with open(json_path, 'w', encoding='utf-8') as json_file:
    json.dump(json_data, json_file, ensure_ascii=False, indent=2)

  print(f'Таблица {table_name} успешно преобразована в JSON')

# Функция для обработки каждого WorkBook Excel
def parse_wb(wb_path):
  wb_handler = openpyxl.load_workbook(wb_path, keep_links=True, keep_vba=True)

  # Открываем файл для хранения локализованного текста
  csv_locale_file_path = os.path.join(output_dir, f'locale.csv')
  with open(csv_locale_file_path, mode='w', newline='', encoding='utf-8') as locale_file:
    csv_locale_writer = csv.writer(locale_file)
    csv_locale_writer.writerow(["id", "label", "column", "name"])

    # Обрабатываем каждую таблицу в файле
    for wb_sheet_name in wb_handler.sheetnames:
      # Игнорируем листы, начинающиеся с символа '@'
      if wb_sheet_name.startswith('@'):
        print(f'Лист {wb_sheet_name} пропущен (начинается с символа @)')
        continue

      wb_sheet_handler = wb_handler[wb_sheet_name]
      for table_name, table in wb_sheet_handler.tables.items():
        table_range = wb_sheet_handler[table]
        table_data = [[cell.value for cell in row] for row in table_range]
        convert_table_to_csv_json(table_name, table_data, csv_locale_writer)

# Настройка среды выполнения
input_dir = os.getcwd()
output_dir = os.getcwd()

# Сканируем директорию на наличие Excel файлов
for root, dirs, files in os.walk(input_dir):
  for file in files:
    wb_file_path = os.path.join(root, file)

    # Пропуск временных файлов (файлы начинаются с '~' или '.')
    if file.startswith('~') or file.startswith('.'): continue

    # Пропуск файлов с неправильным расширением
    if file.endswith('.xlsx') or file.endswith('.xls'):
      print(f'Обнаружен файл {wb_file_path}')
      parse_wb(wb_file_path)

print("Все файлы обработаны.")
input("Press Enter to exit...")