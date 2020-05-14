Attribute VB_Name = "m000_EntryPoints"
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

    Dim sFolderPath As String
    Dim sFilePath As String
    Dim wkb As Workbook
    Dim sQueryText As String
    Dim sht As Worksheet
    
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
    sFilePath = sFolderPath & Application.PathSeparator & "MetadataWorksheets.txt"
    sQueryText = _
        "let" & vbCr & _
        "    Source = Csv.Document(File.Contents(""" & _
        sFilePath & """" & _
        "),[Delimiter=""|"", Encoding=1252, QuoteStyle=QuoteStyle.None])," & vbCr & _
        "   PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])" & _
        "in " & vbCr & _
        "   PromotedHeaders"
    CreateWorksheetPowerQuery sht, "qry_TempMetadataWorksheets", sQueryText, "tbl_WorksheetMetadata"
    
    'Generate temp sheet containing metadata for list object fields
    Set sht = wkb.Sheets.Add
    sht.Name = "Temp_ListObjectFields"
    sFilePath = sFolderPath & Application.PathSeparator & "ListObjectFields.txt"
    sQueryText = _
        "let" & vbCr & _
        "    Source = Csv.Document(File.Contents(""" & _
        sFilePath & """" & _
        "),[Delimiter=""|"", Encoding=1252, QuoteStyle=QuoteStyle.None])," & vbCr & _
        "   PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])" & _
        "in " & vbCr & _
        "   PromotedHeaders"
    CreateWorksheetPowerQuery sht, "qry_TempListObjectFields", sQueryText, "tbl_ListObjectFields"
    
    
    'Generate temp sheet containing metadata for list object values
    Set sht = wkb.Sheets.Add
    sht.Name = "Temp_ListObjectValues"
    sFilePath = sFolderPath & Application.PathSeparator & "ListObjectFieldValues.txt"
    sQueryText = _
        "let" & vbCr & _
        "    Source = Csv.Document(File.Contents(""" & _
        sFilePath & """" & _
        "),[Delimiter=""|"", Encoding=1252, QuoteStyle=QuoteStyle.None])," & vbCr & _
        "   PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])" & _
        "in " & vbCr & _
        "   PromotedHeaders"
    CreateWorksheetPowerQuery sht, "qry_TempListObjectValues", sQueryText, "tbl_ListObjectValues"
    
    
    CreateWorksheets wkb
    
    'InsertIndexPage ActiveWorkbook
    
        
    'Cleanup
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    

End Sub


Public Sub ExportVBAModules()
'Saves active workbook and exports file to VBA_Code subfolder in path of active workbook
' *****IMPORTANT NOTE****
' Any existing files in this subfolder will be deleted

    Dim sExportPath As String
    Dim sExportFileName As String
    Dim bExport As Boolean
    Dim sFileName As String
    Dim cmpComponent As VBIDE.VBComponent

    
    'ActiveWorkbook.Save
    sExportPath = ActiveWorkbook.Path & Application.PathSeparator & "VBA_Code"
    
    If NumberOfFilesInFolder(sExportPath) <> 0 Then
        MsgBox ("Please ensure VBA subfolder is empty ...exiting")
        Exit Sub
    End If

    
    
    On Error Resume Next
        MkDir sExportPath
        Kill sExportPath & "\*.*"
    On Error GoTo 0

    If ActiveWorkbook.VBProject.Protection = 1 Then
        MsgBox "The VBA in this workbook is protected," & _
            "not possible to export the code"
        Exit Sub
    End If
    
    For Each cmpComponent In ThisWorkbook.VBProject.VBComponents
        
        bExport = True
        sFileName = cmpComponent.Name

        'Set filename
        Select Case cmpComponent.Type
            Case vbext_ct_ClassModule
                sFileName = cmpComponent.Name & ".cls"
            Case vbext_ct_MSForm
                sFileName = cmpComponent.Name & ".frm"
            Case vbext_ct_StdModule
                sFileName = cmpComponent.Name & ".bas"
            Case vbext_ct_Document
                ' This is a worksheet or workbook object - don't export.
                bExport = False
        End Select
        
        If bExport Then
            sExportFileName = sExportPath & Application.PathSeparator & sFileName
            cmpComponent.Export sExportFileName
        End If
   
    Next cmpComponent

    MsgBox "Code export complete"
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






Private Sub CreateWorksheetPowerQuery( _
    ByVal sht As Worksheet, _
    ByVal sQueryName As String, _
    ByVal sQueryText As String, _
    ByVal sTableName As String)

    
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


Sub CreateWorksheets(ByRef wkb As Workbook)

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
                Set rngForTable = rngForTable.Resize(2, .ListColumns("Number Of Table Columns").DataBodyRange.Cells(i))
                Set lo = sht.ListObjects.Add(SourceType:=xlSrcRange, Source:=rngForTable)
            End If
            
        Next i
    End With




End Sub



Sub FormatTable(lo As ListObject)

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




Function NumberOfFilesInFolder(ByVal sFolderPath As String) As Integer
'Requires refence: Microsoft Scripting Runtime
'This is non-recursive


    Dim oFSO As FileSystemObject
    Dim oFolder As Folder
    
    Set oFSO = New FileSystemObject
    Set oFolder = oFSO.GetFolder(sFolderPath)
    NumberOfFilesInFolder = oFolder.Files.Count


End Function

