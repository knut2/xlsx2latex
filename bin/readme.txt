xlsx2LaTeX
===========
Convert Excel-sheets into LaTeX-tables.

You can take the LaTeX-table and continue with your work.

Some features:
* Formats for numbers, booleans, dates and times can be defined.
* Selection of datasheet is possible.

Usage: xlsx2latex.exe excelfile [OPTIONS]

Options
    -t, --target TARGET              target filename (default STDOUT)
    -s, --sheet SHEET                Source sheet in Excel (default: all)
    -f, --float FLOAT                Format for floats
    -w, --width WIDTH                Default column width (or 'auto')
    -d, --date DATE                  Format for dates
        --time TIME                  Format for times
        --true TRUE                  Format for true
        --false FALSE                Format for false
    -h, --help                       help

==Examples:
Write default data sheet as excel to stdout:
  xlsx2latex my_excel.xlsx
  
Write sheet with name "DATA2" as excel to stdout:
  xlsx2latex my_excel.xlsx -s DATA2

No decimals for all floats (numbers):
  xlsx2latex my_excel.xlsx -f %.0f

Change data format (e.g. to "1. January 2015"):
  xlsx2latex my_excel.xlsx -d "%d. %B %Y"

Change format for boolean values:
  xlsx2latex my_excel.xlsx  --true yes --false no

=======
Not supported features:
* xls-files [1]
* deep control on text formats [2]
* Suppress text formats (make no bold...) [3]
* select areas or specific columns for export [3]
* column specific formats [3]
* Detect number format of Excel and use it for the output [4]

remarks:
[1] maybe possible in future - not tested yet.
[2] Not planned in future
[3] Not planned yet, but on the todo list.
[4] Not realized, but wanted

