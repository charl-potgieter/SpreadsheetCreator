# Spreadsheet Creator

> Creates Excel spreadsheet utilising metadata stored in text files


## "Installation"

Download Excel file under releases:
https://github.com/charl-potgieter/SpreadsheetCreator/releases


Alternatively:
- Copy and paste the code per "m000_EntryPoints.bas" into an Excel VBA editor of a new workbook
- Add references for 
    - Microsoft scripting runtime
    - Microsoft Visual Basic For Applications Extensibility 5.3
-  Run the GenerateSpreadsheet sub

- Code can be tested by running and selecting the ExampleSpreadsheetMetadata folder included in this repository.


## About this code

- Creates an Excel spreadsheet from metadata saved in text files

- Below code can be copied and pasted into the Excel VBA editor in a  temporary spreadsheet and then run to create the target spreadsheet

- Target spreadsheet can contain one or more below
    - Multiple worksheets
    - Single list object per sheet
    - Each sheet can contain sheet header
    - Each sheet can contain sheet category
    - Front index sheet with hyperlink (if target workbook consists of more than one  sheet)
    - VBA code modules
    - Power Queries


## Metadata folder structure

The metadata is saved in pipe delimited text files stored in below Folders:

    SpreadsheetMetadata
        |
        |--- PowerQueries (one file per query with .m file extension)
        |
        |--- VBA_Code (one file per VBA module with .bas file extension)
        |
        |--- WorksheetStructure
                |
                |--- ListObjectFields.txt
                |
                |--- ListObjectFieldValues.txt
                |
                |--- ListObjectFormat.txt
                |
                |--- MetadataWorksheets.txt


## File structure

The files under WorksheetStructure folder above are pipe delimited text files

#### ListObjectFields.txt Fields:
 - SheetName
 - ListObjectName
 - ListObjectHeader
 - IsFormula
 - Formula

#### ListObjectFieldValues.txt
- SheetName
- ListObjectName
- ListObjectHeader
- Value

#### ListObjectFormat.txt
- SheetName
- ListObjectName
- ListObjectHeader
- NumberFormat
- FontColour

#### MetadataWorksheets.txt
- Name
- Sheet Category
- Sheet Header
- Table Name
- Number Of Table Columns
- Number of Table Rows
- Table top left cell


## Generating spreadsheet metadata for use with this code
Refer https://github.com/charl-potgieter/SpreadsheetBI

 ## License

### The MIT License
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT) 
