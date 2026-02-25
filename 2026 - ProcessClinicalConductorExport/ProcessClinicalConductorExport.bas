


Sub ProcessClinicalConductorExport()
    

    Dim ws          As Worksheet
    Dim rngTable    As Range
    Dim headerRow   As Range
    Dim colHeader   As Range
    Dim dataRange   As Range
    Dim cell        As Range
    Dim screenCol   As Long
    Dim col         As Long
    Dim lastRow     As Long
    Dim lastCol     As Long
    
    Const HEADER_ROW    As Long = 4
    Const DEBUG_MODE    As Boolean = True   ' Set to False in production
    
    
    Set ws = ActiveSheet
    
    Set rngTable = ws.Range("A4").CurrentRegion
    
    ' We want to start 3 rows down
    ' The range would now include blank rows, so we resize to make rngTable 3 rows smaller
    Set rngTable = rngTable.Offset(3).Resize(rngTable.Rows.Count - 3)
    
    ' If the table is tiny, alert user and exit
    If rngTable.Rows.Count < 2 Then
        MsgBox "No data rows found starting from row " & HEADER_ROW, vbExclamation
        GoTo CleanExit
    End If
    
    ' Headers are the first row of rngTable.  This will be in the row number of HEADER_ROW.
    Set headerRow = rngTable.Rows(1)
    
    ' dataRange is everything below headers
    Set dataRange = rngTable.Offset(1, 0).Resize(rngTable.Rows.Count - 1)
    
    ' Set last row and col, used to scan over table
    lastRow = dataRange.Rows.Count + HEADER_ROW       '  last data row
    lastCol = headerRow.Columns.Count + 0           '  last data column number (A=1, B=2, etc)

            
    If DEBUG_MODE Then
        Debug.Print "Initial table: " & rngTable.Address
        Debug.Print "Headers:      " & headerRow.Address
        Debug.Print "Data rows:    " & dataRange.Address
        Debug.Print "Last row:     " & lastRow
        Debug.Print "Last col:     " & lastCol
    End If
    
    ' Looping backwards, because we are deleting columns
    For col = lastCol To 2 Step -1
    
        ' Gets the column heading
        Set colHeader = headerRow.Cells(1, col)
    
        ' Check if column head contains "(Monitored)" or is equal to "Prescreen"
        If InStr(colHeader.Value, "(Monitored)") > 0 Or colHeader.Value = "Prescreen" Then
        
            ' Delete Monitored column immediately
            If DEBUG_MODE Then Debug.Print "Deleting Monitored column:  " & colHeader.Address & " " & colHeader.Value
            ws.Columns(col).Delete Shift:=xlToLeft
    
        ' Check if column heading contains "(Status)"
        ' This is to clear dates in Attribute columns
        ElseIf InStr(colHeader.Value, "(Status)") Then
            If DEBUG_MODE Then Debug.Print "Processing Status column:  " & colHeader.Address & " " & colHeader.Value
        
            ' Look down this column in the data rows
            For Each cell In ws.Range(ws.Cells(5, col), ws.Cells(lastRow, col))
                
                ' Check if cell value is "Completed"
                If Trim(cell.Value) <> "Completed" Then
                    ' Clear the cell to the left (previous column, same row)
                        cell.Offset(0, -1).ClearContents
                End If
                
            Next cell
            
            '  Delete Status column, AFTER processing
            If DEBUG_MODE Then Debug.Print "Deleting Status column:  " & colHeader.Address & " " & colHeader.Value
            ws.Columns(col).Delete Shift:=xlToLeft
        
        ' Checkif column heading is Status
        ElseIf colHeader.Value = "Status" Then
            ' Delete rows where value is "Non-Qualified"
            
            If DEBUG_MODE Then Debug.Print "Deleting rows with Non-Qualified in Status column:  " & colHeader.Address
            
            Dim r As Long
            For r = lastRow To 5 Step -1
                If Trim(ws.Cells(r, col).Value) = "Non-Qualified" Then
                    ws.Rows(r).Delete Shift:=xlUp
                End If
            Next r
            
        End If
        
    Next col
    
    
    ' Sort by Screen# column
    
    ' Try to find "Screen#"
    colHeader = headerRow.Find("Screen#", LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=True)
    
    ' If can't find, warn user and exit gracefully
    If colHeader Is Nothing Then
        MsgBox "Column 'Screen#' not found in header row.", vbExclamation
        GoTo CleanExit
    End If
    
    ' Found "Screen#" in headerRow, so get column number for sorting
    screenCol = colHeader.Column
    
    If DEBUG_MODE Then Debug.Print "Sorting by Screen# in column " & screenCol

    rngTable.Sort Key1:=rngTable.Columns(screenCol - rngTable.Column + 1), _
              Order1:=xlAscending, Header:=xlYes
    
    
CleanExit:
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    MsgBox "Processing complete.", vbInformation
    
End Sub

