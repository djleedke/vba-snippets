Attribute VB_Name = "TableFunctions"

Sub AddDataRow(worksheetName As String, tableName As String, values() As Variant)

    '   Adds a row of data to the specified table on the specified sheet given an array of values
    '
    '   Arguments:
    '       worksheetName: The name of the worksheet the table is on.
    '       tableName: The name of the table.
    '       values(): The array of values that will be added to the table.

    Dim sheet As Worksheet
    Dim table As ListObject
    Dim col As Integer
    Dim lastRow As Range

    Set sheet = ThisWorkbook.Worksheets(worksheetName)
    Set table = sheet.ListObjects.Item(tableName)

    'First check if the last row is empty; if not, add a row
    If table.ListRows.Count > 0 Then
        Set lastRow = table.ListRows(table.ListRows.Count).Range
        For col = 1 To lastRow.Columns.Count
            If Trim(CStr(lastRow.Cells(1, col).Value)) <> "" Then
                table.ListRows.Add
                Exit For
            End If
        Next col
    Else
        table.ListRows.Add
    End If

    'Iterate through the last row and populate it with the entries from values()
    Set lastRow = table.ListRows(table.ListRows.Count).Range
    For col = 1 To lastRow.Columns.Count
        If col <= UBound(values) + 1 Then lastRow.Cells(1, col) = values(col - 1)
    Next col
    
End Sub
