# Spreadsheet Creator

> Creates Excel spreadsheet utilising metadata stored in text files


## "Installation"

- TBA
 - 'Mention something about download'


## What this code does

- Creates an Excel spreadsheet from metada saved in text files

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
        |   ------- PowerQueries (one file per query with .m file extension)
        |
        |   ------- VBA_Code (one file per VBA module with .bas file extension)
        |
        |   ------- WorksheetStructure
                |
                |   ------- ListObjectFields.txt
                |
                |   ------- ListObjectFieldValues.txt
                |
                |   ------- ListObjectFormat.txt
                |
                |   ------- MetadataWorksheets.txt


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

 ## License

### The MIT License
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT) 
