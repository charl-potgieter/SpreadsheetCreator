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
    Dim wkb As Workbook
    
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
    
    InsertIndexPage ActiveWorkbook
    
        
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



Function SheetLevelRangeNameExists(sht As Worksheet, ByRef sRangeName As String)
'Returns TRUE if sheet level scoped range name exists

    Dim sTest As String
    
    On Error Resume Next
    sTest = sht.Names(sRangeName).Name
    SheetLevelRangeNameExists = (Err.Number = 0)
    On Error GoTo 0


End Function


Sub FormatSheet(ByRef sht As Worksheet)
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



Sub InsertIndexPage(ByRef wkb As Workbook)

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
