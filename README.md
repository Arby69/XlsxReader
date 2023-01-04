# XlsxReader
Library for fast and lightweight access to Open Document XLSX (Microsoft Excel) files.

XlsxReader is a .NET Standard 2.0 project without any dependencies (but .NET Standard 2.0 of course),
that lets you open and read xlsx files easy and fast, to access its table data directly or as a `DataSet`
(Workbook) or `DataTable` (Worksheet).

It is not designed for editing purposes, thus you only get a reader, but cannot change the data to write it 
back into the xlsx file.

If you have any ideas about what this library should also support, feel free to tell me your thoughts or 
change the code yourself and do a pull request.

## Usage
The main class is `Workbook`. Use the constructor `Workbook(string filepath)` to directly open your xlsx 
file and generate the appropriate Workbook child elements you want to access.

* `Workbook`
  * `Worksheet`
    * `Row`
      * `Cell`
      
All row and column number indices are 1-based, so a cell of `H4` in Excel is accessed as row 4 and 
column 8 (H = 8th letter).

All lists (of Worksheets, of Rows or Cells) provide an Enumerator, so you are able to browse the data 
with `foreach`.

To access a cell directly use row and column numbers, or you may use the "A1" notation. Make sure 
your adress string consists only of ASCII letters and digits 0 to 9, spaces or other characters 
are not allowed.
