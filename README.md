# cas-data-designer
A set of auxiliary utilities for game design and development.

## Core Settings

- ai_counter_locale_table - stores the autoincrement counter for strings with localized textж
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

1) Open WPS Office Spreadsheets.
2) Click the Developer Tab (depending on your version of WPS).
3) Click “Visual Basic Editor”
4) Modules -> right click -> import 'add-row.bas' and 'toggle-lid-columns.bas'
5) Or Open 

6) After creating a macro, go to File → Options (or WPS Settings).
7) In the window that opens, select the “Customize Ribbon” or “Customize Toolbar” section.
8) Find or create a new tab or group on the ribbon:
9) Click “Create Tab” and name it (for example, “Game Data Designer”).
10) Add 'Add Row' and 'ToggleVisibility' buttons