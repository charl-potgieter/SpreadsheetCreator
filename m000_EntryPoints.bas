Option Explicit

'-------------------------------------------------------------------------------------------------------------------
'       General notes
'-------------------------------------------------------------------------------------------------------------------
'
'
'   References required
'       - Microsoft scripting runtime
'       - Microsoft Visual Basic For Applications Extensibility 5.3
'
'
'   All workings contained in one module to enable easy copy and paste setup



'-------------------------------------------------------------------------------------------------------------------
'       Entry Points
'-------------------------------------------------------------------------------------------------------------------


Sub GenerateSpreadsheet()
'Generates spreadsheet from metadata stored in text files in selected folder


    Dim sFolderPath As String
    Dim sFilePath As String
    Dim wkb As Workbook
    Dim sQueryText As String
    Dim sht As Worksheet
    Dim lo As ListObject
    Dim cn As WorkbookConnection
    
    'Setup
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    
    'Get folder containing metadata
    sFolderPath = GetFolder
    If sFolderPath = "" Then
        Exit Sub
    End If
    
    'Create new workbook with only one sheet
    Set wkb = Application.Workbooks.Add
    Do While wkb.Sheets.Count > 1
        wkb.Sheets(1).Delete
    Loop
    
    
    'Generate temp sheet containing metadata for worksheets
    Set sht = wkb.Sheets(1)
    sht.Name = "Temp_WorksheetMetadata"
    sFilePath = sFolderPath & Application.PathSeparator & _
        "WorksheetStructure" & Application.PathSeparator & "MetadataWorksheets.txt"
        
    sQueryText = _
        "let" & vbCr & _
        "    Source = Csv.Document(File.Contents(""" & _
        sFilePath & """" & _
        "),[Delimiter=""|"", Encoding=1252, QuoteStyle=QuoteStyle.None])," & vbCr & _
        "   PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])" & _
        "in " & vbCr & _
        "   PromotedHeaders"
    CreatePowerQuery sht, "qry_TempMetadataWorksheets", sQueryText, "tbl_WorksheetMetadata"
    
    'Generate temp sheet containing metadata for list object fields
    Set sht = wkb.Sheets.Add
    sht.Name = "Temp_ListObjectFields"
    sFilePath = sFolderPath & Application.PathSeparator & _
        "WorksheetStructure" & Application.PathSeparator & "ListObjectFields.txt"
    
    sQueryText = _
        "let" & vbCr & _
        "    Source = Csv.Document(File.Contents(""" & _
        sFilePath & """" & _
        "),[Delimiter=""|"", Encoding=1252, QuoteStyle=QuoteStyle.None])," & vbCr & _
        "   PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])" & _
        "in " & vbCr & _
        "   PromotedHeaders"
    CreatePowerQuery sht, "qry_TempListObjectFields", sQueryText, "tbl_ListObjectFields"
    
    
    'Generate temp sheet containing metadata for list object values
    Set sht = wkb.Sheets.Add
    sht.Name = "Temp_ListObjectValues"
    sFilePath = sFolderPath & Application.PathSeparator & _
        "WorksheetStructure" & Application.PathSeparator & "ListObjectFieldValues.txt"
    
    sQueryText = _
        "let" & vbCr & _
        "    Source = Csv.Document(File.Contents(""" & _
        sFilePath & """" & _
        "),[Delimiter=""|"", Encoding=1252, QuoteStyle=QuoteStyle.None])," & vbCr & _
        "   PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])" & _
        "in " & vbCr & _
        "   PromotedHeaders"
    CreatePowerQuery sht, "qry_TempListObjectValues", sQueryText, "tbl_ListObjectValues"
    
    
    'Generate temp sheet containing metadata for list object format
    Set sht = wkb.Sheets.Add
    sht.Name = "Temp_ListObjectFormats"
    sFilePath = sFolderPath & Application.PathSeparator & _
        "WorksheetStructure" & Application.PathSeparator & "ListObjectFormat.txt"
    
    sQueryText = _
        "let" & vbCr & _
        "    Source = Csv.Document(File.Contents(""" & _
        sFilePath & """" & _
        "),[Delimiter=""|"", Encoding=1252, QuoteStyle=QuoteStyle.None])," & vbCr & _
        "   PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])" & _
        "in " & vbCr & _
        "   PromotedHeaders"
    CreatePowerQuery sht, "qry_TempListObjectFormats", sQueryText, "tbl_ListObjectFormats"
    
    
    CreateWorksheets wkb
    PopulateListFieldNamesAndFormulas wkb
    PopulateListObjectValues wkb
    SetListObjectFormats wkb
        
    
    'Delete temp sheets
    wkb.Sheets("Temp_ListObjectFormats").Delete
    wkb.Sheets("Temp_ListObjectValues").Delete
    wkb.Sheets("Temp_ListObjectFields").Delete
    wkb.Sheets("Temp_WorksheetMetadata").Delete
    
    'Delete temp queries
    wkb.Queries("qry_TempListObjectFormats").Delete
    wkb.Queries("qry_TempListObjectValues").Delete
    wkb.Queries("qry_TempListObjectFields").Delete
    wkb.Queries("qry_TempMetadataWorksheets").Delete
    
    'Delete any worbook connection (seems like one is created for each of the temp queries)
    For Each cn In wkb.Connections
        cn.Delete
    Next cn
    
    
    'Set table styles and Freeze panes
    For Each sht In wkb.Sheets
        sht.Select
        For Each lo In sht.ListObjects
            sht.Rows(lo.DataBodyRange.Cells(1).Row).Select
            ActiveWindow.FreezePanes = True
            lo.DataBodyRange.Cells(1).Select
            FormatTable lo
        Next lo
    Next sht
    
    'Import power queries
    ImportPowerQueriesInFolder sFolderPath & Application.PathSeparator & "PowerQueries", True
    
    'Import VBA code
    ImportVBAModules wkb, sFolderPath & Application.PathSeparator & "VBA_Code"
    
    'Create index tab only if more than one sheet exists in wkb
    If wkb.Sheets.Count > 1 Then
        InsertIndexPage wkb
    End If
    
        
    'Cleanup
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    

End Sub




'-------------------------------------------------------------------------------------------------------------------
'       General
'-------------------------------------------------------------------------------------------------------------------



Private Function SheetLevelRangeNameExists(sht As Worksheet, ByRef sRangeName As String)
'Returns TRUE if sheet level scoped range name exists

    Dim sTest As String
    
    On Error Resume Next
    sTest = sht.Names(sRangeName).Name
    SheetLevelRangeNameExists = (Err.Number = 0)
    On Error GoTo 0


End Function


Private Sub FormatSheet(ByRef sht As Worksheet)
'Applies my preferred sheet formattting

    sht.Activate
    
    sht.Cells.Font.Name = "Calibri"
    sht.Cells.Font.Size = 11
    
    sht.Range("A1").Font.Color = RGB(170, 170, 170)
    sht.Range("A1").Font.Size = 8
    
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 80
    sht.DisplayPageBreaks = False
    sht.Columns("A:A").ColumnWidth = 4
    
    If SheetLevelRangeNameExists(sht, "SheetHeading") Then
        sht.Names("SheetHeading").Delete
    End If
    sht.Names.Add Name:="SheetHeading", RefersTo:="=$B$2"
    
    If SheetLevelRangeNameExists(sht, "SheetCategory") Then
        sht.Names("SheetCategory").Delete
    End If
    sht.Names.Add Name:="SheetCategory", RefersTo:="=$A$1"
    
    With sht.Range("SheetHeading")
        If .Value = "" Then
            .Value = "Heading"
        End If
        .Font.Bold = True
        .Font.Size = 16
    End With

End Sub



Private Sub InsertIndexPage(ByRef wkb As Workbook)
'Inserts index page with hyperlinks to subsequent sheets

    Dim sht As Worksheet
    Dim shtIndex As Worksheet
    Dim i As Double
    Dim sPreviousReportCategory As String
    Dim sReportCategory As String
    Dim sReportName As String
    Dim rngCategoryCol As Range
    Dim rngReportCol As Range
    Dim rngSheetNameCol As Range
    Dim rngShowRange As Range
    
    'Delete any previous index sheet and create a new one
    On Error Resume Next
    wkb.Sheets("Index").Delete
    On Error GoTo 0
    Set shtIndex = wkb.Sheets.Add(Before:=ActiveWorkbook.Sheets(1))
    FormatSheet shtIndex
    
    wkb.Activate
    shtIndex.Activate
    
    With shtIndex
    
        .Name = "Index"
        .Range("A:A").Insert Shift:=xlToRight
        .Range("A:A").EntireColumn.Hidden = True
        .Range("C2") = "Index"
        .Range("D5").Font.Bold = True
        .Columns("D:D").ColumnWidth = 100
        .Rows("4:4").Select
        ActiveWindow.FreezePanes = True
        
        Set rngSheetNameCol = .Columns("A")
        Set rngCategoryCol = .Columns("C")
        Set rngReportCol = .Columns("D")
       
        sPreviousReportCategory = ""
        i = 2
        
        
        For Each sht In wkb.Worksheets
        
            sReportCategory = sht.Range("A1")
            sReportName = sht.Range("B2")
            
            If (sReportCategory <> "" And sReportName <> "") And (sht.Name <> "Index") And (sht.Visible = xlSheetVisible) Then
            
                'Create return to Index links
                sht.Hyperlinks.Add _
                    Anchor:=sht.Range("B3"), _
                    Address:="", _
                    SubAddress:="Index!A1", _
                    TextToDisplay:="<Return to Index>"
                    
                'Write the report category headers
                If sReportCategory <> sPreviousReportCategory Then
                    i = i + 3
                    rngCategoryCol.Cells(i) = sReportCategory
                    rngCategoryCol.Cells(i).Font.Bold = True
                    sPreviousReportCategory = sReportCategory
                End If
    
                i = i + 2
                rngReportCol.Cells(i) = sReportName
                rngSheetNameCol.Cells(i) = sht.Name
                
                ActiveSheet.Hyperlinks.Add _
                    Anchor:=rngReportCol.Cells(i), _
                    Address:="", _
                    SubAddress:="'" & sht.Name & "'" & "!B$4"
                    
            End If
            
        Next sht
        
        .Range("C3").Select
        
    End With


End Sub




Private Sub CreatePowerQuery( _
    ByVal sht As Worksheet, _
    ByVal sQueryName As String, _
    ByVal sQueryText As String, _
    ByVal sTableName As String)

'Creates power query and loads as a table on sht
    
        
    Dim lo As ListObject
    
    sht.Parent.Queries.Add sQueryName, sQueryText
        
    'Output the Power Query to a worksheet table
    Set lo = sht.ListObjects.Add( _
        SourceType:=0, _
        Source:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & sQueryName & ";Extended Properties=""""", _
        Destination:=Range("$A$1"))
        
    lo.Name = sTableName
    
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & sQueryName & "]")
        .Refresh BackgroundQuery:=False
    End With
    
    
        
End Sub

Private Sub ImportSinglePowerQuery(ByVal sQueryPath As String, ByVal sQueryName As String, wkb As Workbook)

    Dim sQueryText As String
    
    sQueryText = ReadTextFileIntoString(sQueryPath)
    wkb.Queries.Add sQueryName, sQueryText

End Sub



Private Sub ImportPowerQueriesInFolder(ByVal sFolderPath As String, ByVal bRecursive As Boolean)
'Reference: Microsoft Scripting Runtime
    
    Dim FileItems() As Scripting.File
    Dim FileItem
    Dim sQueryName As String
    
    FileItemsInFolder sFolderPath, bRecursive, FileItems
    
    For Each FileItem In FileItems
        sQueryName = Left(FileItem.Name, Len(FileItem.Name) - 2)
        ImportSinglePowerQuery FileItem.Path, sQueryName, ActiveWorkbook
    Next FileItem


End Sub




Private Sub CreateWorksheets(ByRef wkb As Workbook)
'Creates worksheets based on data stored in the listobject on the Temp_WorksheetMetadata tab of wkb workkbook


    Dim i As Long
    Dim loSheetMetadata As ListObject
    Dim lo As ListObject
    Dim sht As Worksheet
    Dim rngForTable As Range
    
    Set loSheetMetadata = wkb.Sheets("Temp_WorksheetMetadata").ListObjects("tbl_WorksheetMetadata")
    
    With loSheetMetadata
        For i = 1 To .DataBodyRange.Rows.Count
            Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
            FormatSheet sht
            sht.Name = .ListColumns("Name").DataBodyRange.Cells(i)
            sht.Names("SheetCategory").RefersToRange = .ListColumns("Sheet Category").DataBodyRange.Cells(i)
            sht.Names("SheetHeading").RefersToRange = .ListColumns("Sheet Header").DataBodyRange.Cells(i)
            
            If .ListColumns("Table Name").DataBodyRange.Cells(i) <> "" Then
                Set rngForTable = sht.Range(.ListColumns("Table top left cell").DataBodyRange.Cells(i))
                Set rngForTable = rngForTable.Resize(.ListColumns("Number Of Table Rows").DataBodyRange.Cells(i) - 1, .ListColumns("Number Of Table Columns").DataBodyRange.Cells(i))
                Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=rngForTable)
                lo.Name = .ListColumns("Table Name").DataBodyRange.Cells(i)
            End If
            
        Next i
    End With




End Sub


Private Sub PopulateListFieldNamesAndFormulas(ByRef wkb As Workbook)
'Records listobject field names and formulas in wkb based on metadata stored in
'wkb.Sheets("Temp_ListObjectFields").ListObjects("tbl_ListObjectFields")

    Dim loFieldDetails As ListObject
    Dim loTargetListObj As ListObject
    Dim i As Long
    Dim j As Long
    Dim sSheetName As String
    Dim sListObjName As String
    Dim sListObjHeader As String
    Dim bIsFormula As Boolean
    Dim sFormula As String

    Set loFieldDetails = wkb.Sheets("Temp_ListObjectFields").ListObjects("tbl_ListObjectFields")
    
    'No ListObject field details.  Noting to do.  Sub exited to prevent error
    'caused by referencing the list databodyrange
    If loFieldDetails.DataBodyRange Is Nothing Then
        Exit Sub
    End If
    
    
    With loFieldDetails
        For i = 1 To .DataBodyRange.Rows.Count
            sSheetName = .ListColumns("SheetName").DataBodyRange.Cells(i)
            sListObjName = .ListColumns("ListObjectName").DataBodyRange.Cells(i)
            sListObjHeader = .ListColumns("ListObjectHeader").DataBodyRange.Cells(i)
            bIsFormula = CBool(.ListColumns("isFormula").DataBodyRange.Cells(i))
            sFormula = .ListColumns("Formula").DataBodyRange.Cells(i)
            
            'Increment j as header col counter if writing to the table name as previous iteration
            If i = 1 Then
                j = 1
            ElseIf sListObjName = .ListColumns("ListObjectName").DataBodyRange.Cells(i - 1) Then
                j = j + 1
            Else
                j = 1
            End If
            
            Set loTargetListObj = wkb.Worksheets(sSheetName).ListObjects(sListObjName)
            loTargetListObj.HeaderRowRange.Cells(j) = sListObjHeader
            
            If bIsFormula Then
                loTargetListObj.ListColumns(sListObjHeader).DataBodyRange.Formula = sFormula
            End If
            
            
        Next i
    End With
    
    
    
    

End Sub




Private Sub PopulateListObjectValues(ByRef wkb As Workbook)
'Populates listobject values in wkb based on values stored in
'wkb.Sheets("Temp_ListObjectValues").ListObjects("tbl_ListObjectValues")


    Dim loListObjValues As ListObject
    Dim i As Long
    Dim j As Long
    Dim sTargetSheetName As String
    Dim sTargetListObjName As String
    Dim sTargetListColName As String
    Dim vTargetValue As Variant
    
    
    Set loListObjValues = wkb.Sheets("Temp_ListObjectValues").ListObjects("tbl_ListObjectValues")

    'No ListObject field details.  Noting to do.  Sub exited to prevent error
    'caused by referencing the list databodyrange
    If loListObjValues.DataBodyRange Is Nothing Then
        Exit Sub
    End If


    With loListObjValues
        For i = 1 To .DataBodyRange.Rows.Count
            
            sTargetSheetName = .ListColumns("SheetName").DataBodyRange.Cells(i)
            sTargetListObjName = .ListColumns("ListObjectName").DataBodyRange.Cells(i)
            sTargetListColName = .ListColumns("ListObjectHeader").DataBodyRange.Cells(i)
            vTargetValue = .ListColumns("Value").DataBodyRange.Cells(i)
            
            'Increment j as row to write counter if writing the same column and table name as previous iteration
            If i = 1 Then
                j = 1
            ElseIf sTargetListObjName = .ListColumns("ListObjectName").DataBodyRange.Cells(i - 1) And _
                sTargetListColName = .ListColumns("ListObjectHeader").DataBodyRange.Cells(i - 1) Then
                    j = j + 1
            Else
                j = 1
            End If
            
            wkb.Sheets(sTargetSheetName).ListObjects(sTargetListObjName).ListColumns(sTargetListColName).DataBodyRange.Cells(j) = vTargetValue
            
            
        Next i
    End With



End Sub


Private Sub SetListObjectFormats(ByRef wkb As Workbook)
'Sets font colour and number format of listobject columns in wkb based on metadata in
'wkb.Sheets("Temp_ListObjectFormats").ListObjects("tbl_ListObjectFormats")

    Dim loListObjFormats As ListObject
    Dim i As Long
    Dim sTargetSheetName As String
    Dim sTargetListObjName As String
    Dim sTargetListColName As String
    Dim sNumberFormat As String
    Dim lFontColour As Long
    
    
    Set loListObjFormats = wkb.Sheets("Temp_ListObjectFormats").ListObjects("tbl_ListObjectFormats")

    'No ListObject field details.  Noting to do.  Sub exited to prevent error
    'caused by referencing the list databodyrange
    If loListObjFormats.DataBodyRange Is Nothing Then
        Exit Sub
    End If

    With loListObjFormats
        For i = 1 To .DataBodyRange.Rows.Count
            
            sTargetSheetName = .ListColumns("SheetName").DataBodyRange.Cells(i)
            sTargetListObjName = .ListColumns("ListObjectName").DataBodyRange.Cells(i)
            sTargetListColName = .ListColumns("ListObjectHeader").DataBodyRange.Cells(i)
            sNumberFormat = .ListColumns("NumberFormat").DataBodyRange.Cells(i)
            lFontColour = .ListColumns("FontColour").DataBodyRange.Cells(i)
            
            wkb.Sheets(sTargetSheetName).ListObjects(sTargetListObjName).ListColumns(sTargetListColName).DataBodyRange.NumberFormat = sNumberFormat
            wkb.Sheets(sTargetSheetName).ListObjects(sTargetListObjName).ListColumns(sTargetListColName).DataBodyRange.Font.Color = lFontColour
            
        Next i
    End With



End Sub


Private Sub FormatTable(lo As ListObject)
'Applies preferred listobject formatting

    Dim sty As TableStyle
    Dim wkb As Workbook
    
    Set wkb = lo.Parent.Parent
    
    On Error Resume Next
    wkb.TableStyles.Add ("SpreadsheetBiStyle")
    On Error GoTo 0
    Set sty = wkb.TableStyles("SpreadsheetBiStyle")
    
    'Set Header Format
    With sty.TableStyleElements(xlHeaderRow)
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Borders.Item(xlEdgeTop).LineStyle = xlSolid
        .Borders.Item(xlEdgeTop).Weight = xlMedium
        .Borders.Item(xlEdgeBottom).LineStyle = xlSolid
        .Borders.Item(xlEdgeBottom).Weight = xlMedium
    End With

    'Set row stripe format
    sty.TableStyleElements(xlRowStripe1).Interior.Color = RGB(217, 217, 217)
    sty.TableStyleElements(xlRowStripe2).Interior.Color = RGB(255, 255, 255)
    
    'Set whole table bottom edge format
    sty.TableStyleElements(xlWholeTable).Borders.Item(xlEdgeBottom).LineStyle = xlSolid
    sty.TableStyleElements(xlWholeTable).Borders.Item(xlEdgeBottom).Weight = xlMedium

    
    'Apply custom style and set other attributes
    lo.TableStyle = "SpreadsheetBiStyle"
    With lo.HeaderRowRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
    End With
    
    lo.DataBodyRange.EntireColumn.AutoFit


End Sub


Private Function ArrayIsDimensioned(Arr As Variant) As Boolean

    Dim b As Boolean
    
    On Error Resume Next
    ArrayIsDimensioned = (UBound(Arr, 1)) >= 0 And UBound(Arr) >= LBound(Arr)
    If Err.Number <> 0 Then ArrayIsDimensioned = False

End Function


Private Sub ImportVBAModules(ByRef wkb As Workbook, ByVal sFolder As String)
'Imports VBA code sFolder


    Dim objFSO As Scripting.FileSystemObject
    Dim objFile As Scripting.File
    Dim sTargetWorkbook As String
    Dim sImportPath As String
    Dim zFileName As String
    Dim cmpComponents As VBIDE.VBComponents

    If wkb.Name = ThisWorkbook.Name Then
        MsgBox "Select another destination workbook" & _
        "Not possible to import in this workbook "
        Exit Sub
    End If
    
    If wkb.VBProject.Protection = 1 Then
    MsgBox "The VBA in this workbook is protected," & _
        "not possible to Import the code"
    Exit Sub
    End If

        
    Set objFSO = New Scripting.FileSystemObject
    If objFSO.GetFolder(sFolder).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If


    Set cmpComponents = wkb.VBProject.VBComponents
    
    For Each objFile In objFSO.GetFolder(sFolder).Files
    
        If (objFSO.GetExtensionName(objFile.Name) = "cls") Or _
            (objFSO.GetExtensionName(objFile.Name) = "frm") Or _
            (objFSO.GetExtensionName(objFile.Name) = "bas") Then
            cmpComponents.Import objFile.Path
        End If
        
    Next objFile
    
End Sub



'-------------------------------------------------------------------------------------------------------------------
'       File utilities
'-------------------------------------------------------------------------------------------------------------------


Private Function GetFolder() As String
'Returns the results of a user folder picker

    Dim fldr As FileDialog
    Dim sItem As String
    
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a folder"
        .AllowMultiSelect = False
        .InitialFileName = ActiveWorkbook.Path
        If .Show = -1 Then
            GetFolder = .SelectedItems(1)
        End If
    End With
    
    Set fldr = Nothing


End Function




Private Function NumberOfFilesInFolder(ByVal sFolderPath As String) As Integer
'Requires refence: Microsoft Scripting Runtime
'This is non-recursive


    Dim oFSO As FileSystemObject
    Dim oFolder As Folder
    
    Set oFSO = New FileSystemObject
    Set oFolder = oFSO.GetFolder(sFolderPath)
    NumberOfFilesInFolder = oFolder.Files.Count


End Function



Private Sub FileItemsInFolder(ByVal sFolderPath As String, ByVal bRecursive As Boolean, ByRef FileItems() As Scripting.File)
'Requires refence: Microsoft Scripting Runtime
'Returns an array of files (which can be used to get filename, path etc)
'(Cannot create function due to recursive nature of the code)

    
    Dim FSO As Scripting.FileSystemObject
    Dim SourceFolder As Scripting.Folder
    Dim SubFolder As Scripting.Folder
    Dim FileItem As Scripting.File
    
    Set FSO = New Scripting.FileSystemObject
    Set SourceFolder = FSO.GetFolder(sFolderPath)
    
    For Each FileItem In SourceFolder.Files
    
        If Not ArrayIsDimensioned(FileItems) Then
            ReDim FileItems(0)
        Else
            ReDim Preserve FileItems(UBound(FileItems) + 1)
        End If
        
        Set FileItems(UBound(FileItems)) = FileItem
        
    Next FileItem
    
    If bRecursive Then
        For Each SubFolder In SourceFolder.SubFolders
            FileItemsInFolder SubFolder.Path, True, FileItems
        Next SubFolder
    End If
    
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set FSO = Nothing
    

End Sub



Private Function ReadTextFileIntoString(sFilePath As String) As String
'Inspired by:
'https://analystcave.com/vba-read-file-vba/

    Dim iFileNo As Integer
    
    'Get first free file number
    iFileNo = FreeFile

    Open sFilePath For Input As #iFileNo
    ReadTextFileIntoString = Input$(LOF(iFileNo), iFileNo)
    Close #iFileNo

End Function
