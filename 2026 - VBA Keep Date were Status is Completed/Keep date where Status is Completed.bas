

Sub Keep_Date_Where_Status_Is_Completed()
    
    Dim ws          As Worksheet
    Dim rngTable    As Range
    Dim headerRow   As Range
    Dim colHeader   As Range
    Dim dataRow     As Range
    Dim cell        As Range
    Dim col         As Long
    Dim lastRow     As Long
    Dim lastCol     As Long
    Dim debugMode   As Boolean
    
    debugMode = True
    
    
    Set ws = ActiveSheet
    
    Set rngTable = ws.Range("A4").CurrentRegion
    
    ' We want to start 3 rows down
    ' The range would now include blank rows, so we resize to make rngTable 3 rows smaller
    Set rngTable = rngTable.Offset(3).Resize(rngTable.Rows.Count - 3)
    
    ' If the table is tiny, alert user and exit
    If rngTable.Rows.Count < 2 Then
        MsgBox "No data rows found below row 4.", vbExclamation
        Exit Sub
    End If
    
    ' Headers are in row 4, which is the first row of rngTable
    Set headerRow = rngTable.Rows(1)
    If debugMode Then Debug.Print "headerRow = " & headerRow.Address
    
    ' dataRow = everything below headers
    Set dataRow = rngTable.Offset(1, 0).Resize(rngTable.Rows.Count - 1)
    If debugMode Then Debug.Print "dataRow = " & dataRow.Address

    
    lastRow = dataRow.Rows.Count + 4        '  last data row
    lastCol = headerRow.Columns.Count + 0   '  last data column number (A=1, B=2, etc)
    If debugMode Then Debug.Print "lastRow = " & lastRow
    If debugMode Then Debug.Print "lastCol = " & lastCol
    
    
    ' Looping backwards, because we are deleting columns
    For col = lastCol To 4 Step -1
    
        Set colHeader = headerRow.Cells(1, col)
    
        ' Check if column head contains "(Monitored)"
        If InStr(colHeader.Value, "(Monitored)") > 0 Then
        
            ' Delete Monitored column immediately
            If debugMode Then Debug.Print "Deleting Monitored column:  " & colHeader.Address & " " & colHeader.Value
            ws.Columns(col).Delete Shift:=xlToLeft
    
        ' Check if column header contains "(Status)"
        ElseIf InStr(colHeader.Value, "(Status)") Then
            If debugMode Then Debug.Print "Processing Status column:  " & colHeader.Address & " " & colHeader.Value
        
            ' Look down this column in the data rows
            For Each cell In ws.Range(ws.Cells(5, col), ws.Cells(lastRow, col))
                
                ' Check if cell value is "Completed"
                If Trim(cell.Value) <> "Completed" Then
                    ' Clear the cell to the left (previous column, same row)
                        cell.Offset(0, -1).ClearContents
                End If
                
            Next cell
            
            '  Delete Status column, AFTER processing
            If debugMode Then Debug.Print "Deleting Status column:  " & colHeader.Address & " " & colHeader.Value
            ws.Columns(col).Delete Shift:=xlToLeft
        
        End If
    Next col
    
End Sub

