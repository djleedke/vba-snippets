Attribute VB_Name = "FileHandling"

Function MoveFile(ByVal currentPath As String, ByVal directoryPath As String, Optional ByVal newFileName) As String

    '   Moves a file from the specified path to the specified folder path. Must
    '   have Microsoft Scripting Runtime enabled in Tools -> References
    '
    '   Arguments:
    '       currentPath: The current path of the file.
    '       directoryPath: The path of the folder the file is being moved into.
    '       newFileName: Optional, the new name of the file if desired, do not include extension.
    '
    '   Returns:
    '       String: The new path of the file.

    Dim FSO As New FileSystemObject
    Dim destinationPath As String
    Set FSO = CreateObject("Scripting.FileSystemObject")
    
    Call MakeDirectory(directoryPath)
    
    If IsMissing(newFileName) = False Then
        destinationPath = directoryPath & newFileName & "." & FSO.GetExtensionName(currentPath)
    Else
        destinationPath = directoryPath & Dir(currentPath)
    End If
    
    If Not FSO.FileExists(destinationPath) Then
        FSO.MoveFile currentPath, destinationPath
        MoveFile = destinationPath
        Exit Function
    Else
        answer = MsgBox("File name already exists, overwrite?", vbQuestion + vbYesNo + vbDefaultButton2, "File Exists")
        
        If answer = vbYes Then
            FSO.CopyFile currentPath, destinationPath, True
            FSO.DeleteFile currentPath
            MoveFile = destinationPath
            Exit Function
        End If
        
        If answer = vbNo Then
            End
        End If

    End If
 
    
    
    
    MoveFile = currentPath

End Function

Function MakeDirectory(ByVal directoryPath As String)
    
    '   Makes a directory at the specified path unless one already exists. Must
    '   have Microsoft Scripting Runtime enabled in Tools -> References
    '
    '   Arguments:
    '       directoryPath: The path of the directory that will be created.

    Dim FSO As New FileSystemObject
    Dim checkPath As String
    
    checkPath = ""
    
    'Checking each directory in the path and making folders if they don't exist
    For Each ele In Split(directoryPath, "\")
    
        checkPath = checkPath & ele & "\"
        
        If Len(Dir(checkPath, vbDirectory)) = 0 Then
            FSO.CreateFolder checkPath
        End If
        
    Next

End Function

Function OpenDialogBoxPDF() As String

    '   Display a dialog box that allows the user to select a single PDF file.
    '
    '   Returns:
    '       String: The path of the file chosen.

    With Application.FileDialog(msoFileDialogFilePicker)

        .AllowMultiSelect = False
        .Filters.Add "PDF", "*.pdf", 1
        If .Show = -1 Then
            OpenDialogBoxPDF = .SelectedItems.Item(1)
        Else
            End
        End If
        
    End With

End Function


Function GetWorkbookByPath(ByVal path) As Excel.Application
    
    '   Gets a workbook object at the provided path.
    '
    '   Arguments:
    '       path: The path of the workbook.
    '
    '   Returns:
    '       workbook: An excel workbook object.

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    Dim file
    file = path
    
    If (file = False) Then
        Exit Function 'No file, we're out
    End If

    'Create the Excel object
    On Error Resume Next
    Set objExcel = CreateObject("Excel.Application")
    If (Err.number <> 0) Then
        MsgBox "Failed to create the excel object."
        Exit Function
    End If
    
    'Open the document as read-only
    On Error Resume Next
        Call objExcel.Workbooks.Open(file, False, True)
    If (Err.number <> 0) Then
        MsgBox "Failed to open the document."
    End If
    
    Set GetWorkbookByPath = objExcel
    
Leave:
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic

End Function


Function GetWorkbookBySelect() As Excel.Application

    '   Gets a workbook object at the provided path by allowing the user to select it.
    '
    '   Returns:
    '       workbook:The selected workbook object.

    Application.ScreenUpdating = False
    Application.Calculation = xlManual
    
    Dim file
    'Opens a prompt to select the file
    ChDir Application.ActiveWorkbook.path
    file = Application.GetOpenFilename("Excel File (*.xlsx), *.xlsx")
    
    If (file = False) Then
        Exit Function 'No file, we're out
    End If

    'Create the Excel object
    On Error Resume Next
    Set objExcel = CreateObject("Excel.Application")
    If (Err.number <> 0) Then
        MsgBox "Failed to create the excel object."
        Exit Function
    End If
    
    'Open the document as read-only
    On Error Resume Next
        Call objExcel.Workbooks.Open(file, False, True)
    If (Err.number <> 0) Then
        MsgBox "Failed to open the document."
    End If
    
    Set GetWorkbookBySelect = objExcel
    
Leave:
    Application.ScreenUpdating = True
    Application.Calculation = xlAutomatic
    
End Function

Function ExportWorksheetAndSaveAsCSV(ByVal exportSheet As Worksheet, ByVal savePath As String, ByVal fileName As String, ByVal askForOverwrite As Boolean)

    '   Exports the designated worksheet as a CSV file to the provided directory.
    '
    '   Arguments:
    '       exportSheet: The sheet to be exported.
    '       savePath: The path of the directory the CSV will be saved to.
    '       fileName: The desired name of the file.
    '       askForOverwrite: True if user would like to be prompted to overwrite, False if not.

    Dim exportWorbook As Workbook
    Set exportWorkbook = Application.Workbooks.Add
        
    exportSheet.Copy Before:=exportWorkbook.Worksheets(exportWorkbook.Worksheets.Count)
    
    If (askForOverwrite = False) Then
        Application.DisplayAlerts = False
    End If

    exportWorkbook.SaveAs fileName:=savePath & "\" & fileName & ".csv", FileFormat:=xlCSV
    
    If (askForOverwrite = False) Then
        Application.DisplayAlerts = True
    End If
    
    exportWorkbook.Close SaveChanges:=False


End Function

Function ExportWorksheetAndSaveAsXLSX(ByVal exportSheet As Worksheet, ByVal savePath As String, ByVal fileName As String, ByVal askForOverwrite As Boolean)
    
    '   Exports the designated worksheet as an .xlsx file to the provided directory.
    '
    '   Arguments:
    '       exportSheet: The sheet to be exported.
    '       savePath: The path of the directory the .xlsx will be saved to.
    '       fileName: The desired name of the file.
    '       askForOverwrite: True if user would like to be prompted to overwrite, False if not.
    
    Dim exportWorkbook As Workbook
    Set exportWorkbook = Application.Workbooks.Add
    
    exportSheet.Copy Before:=exportWorkbook.Worksheets(exportWorkbook.Worksheets.Count)
    
    If (askForOverwrite = False) Then
        Application.DisplayAlerts = False
    End If
    
    exportWorkbook.SaveAs fileName:=savePath & "\" & fileName & ".xlsx", FileFormat:=xlOpenXMLWorkbook
    
    exportWorkbook.Close SaveChanges:=False

End Function

Function ImportTextFile(ByVal filePath As String, ByVal startRange As Range)

    '   Imports the text file at the provided path and places the data starting at the specified range.
    '
    '   Arguments:
    '       filePath: The path of the text file to be imported.
    '       startRange: The range object where the data will start from.
    '
    '   Returns:
    '       boolean: True if file was found, False if not.
    
    On Error GoTo ErrorHandler

    Dim StrLine As String
    Dim FSO As New FileSystemObject
    Dim TSO As Object
    Dim StrLineElements As Variant
    Dim Index As Long
    Dim i As Long
    Dim Delimiter As String
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set TSO = FSO.OpenTextFile(filePath)
 
    Delimiter = vbTab
    Index = 0
 
    Do While TSO.AtEndOfStream = False
       StrLine = TSO.ReadLine
       StrLineElements = Split(StrLine, Delimiter)
       For i = LBound(StrLineElements) To UBound(StrLineElements)
           startRange.Offset(Index, i).Value = StrLineElements(i)
       Next i
       Index = Index + 1
    Loop
 
    TSO.Close
    
Success:
    ImportTextFile = True
    Exit Function
    
ErrorHandler:
    ImportTextFile = False
    Exit Function
 
End Function

