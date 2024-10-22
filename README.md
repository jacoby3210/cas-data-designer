# cas-data-designer
A set of auxiliary utilities for game design and development.

## Core Settings

- ai_counter_locale_table - stores the autoincrement counter for strings with localized text.
- show_lid_columns - store current visibility state for lid columns.

## Main Principles

Table Header:
- ':' separates the data type from the column name;
- lid - id of the localized text string;
- ltext - columns store localized text;
- lid, ltext - when exported, they are dumped into locale.json, (as well as column and table name);
- ref - stores reference to external table; 

## Compile:
- If the sheet name starts with @, the compilation script ignores it;
- Columns with the formula data type are ignored;
- Columns with the ref data type leave only the id;

## Steps for adding macros to the ribbon in WPS Office:

1) Copy 'src\excel\DesignDataPrototype.xlsm' and 'src\python\compile.py' in work folder.

2) Open DesignDataPrototype.xlsm (enable macros)
  - ctrl+shift+z - call macro that switches visibility of columns with data type lid
  - ctrl+shift+x - call macro that adds a new row to the table with data 
    - automatically generates new unique lid 
    - copy the range for editing for columns ref (from first table data row).

3) Create new Table (just copy from @examples)
  - columns with name 'id' are automatically filled with the first available value using the autoincrement counter (from column header).
  - columns with name 'label' are automatically filled with the row id.
  - columns with type 'lid' are automatically filled with the first available value using the autoincrement counter (from settings -> ai_counter_locale_table).
  - other data columns of the new row copy data from the first row.
  - make sure that for columns with data type ref, the first row of the table contains a drop-down list with the required data range.

4) Save as .xlsx (current version not supported .xlsm).
5) Run compile.py which created .json files for future use