Attribute VB_Name = "Module1"
Option Explicit

Sub TestTimesheetEntry()
    Dim entry As New TimesheetEntry
    
    entry.EmployeeName = "Karan Singh"
    entry.HoursLogged = 6.5
    entry.Billable = "Y"
    entry.WorkDate = Date
    
    If entry.Validate() Then
        MsgBox "Entry is VALID"
    Else
        MsgBox "Entry is INVALID"
    End If
End Sub

Sub TestInvalidEntry()
    Dim entry As New TimesheetEntry
    
    entry.EmployeeName = ""
    entry.HoursLogged = 30
    entry.Billable = "Maybe"
    entry.WorkDate = Date
    
    If entry.Validate() Then
        MsgBox "Entry is VALID"
    Else
        MsgBox "Entry is INVALID"
    End If
End Sub

Sub ProcessTimesheets()
    On Error GoTo ErrHandler
    Dim wsRaw As Worksheet
    Dim wsProcessed As Worksheet
    Dim rawData As Variant
    Dim outputData() As Variant
    Dim lastRow As Long
    Dim i As Long
    Dim entry As TimesheetEntry
    Dim validCount As Long
    Dim invalidCount As Long
    
    Set wsRaw = ThisWorkbook.Sheets("Raw Data")
    Set wsProcessed = ThisWorkbook.Sheets("Processed")
    
    ' Find the last row with data
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 1).End(xlUp).Row
    
    ' Read the ENTIRE range into an array in one operation (fast)
    rawData = wsRaw.Range("A2:E" & lastRow).Value
    
    ' Size the output array: same rows, one extra column for validation result
    ReDim outputData(1 To UBound(rawData, 1), 1 To 6)
    
    validCount = 0
    invalidCount = 0
    
    ' Loop through the array IN MEMORY (not touching the worksheet each time)
    For i = 1 To UBound(rawData, 1)
        Set entry = New TimesheetEntry
        entry.EmployeeName = rawData(i, 1)
        entry.ProjectName = rawData(i, 2)
        entry.HoursLogged = rawData(i, 3)
        entry.Billable = rawData(i, 4)
        entry.WorkDate = rawData(i, 5)
        
        outputData(i, 1) = rawData(i, 1)
        outputData(i, 2) = rawData(i, 2)
        outputData(i, 3) = rawData(i, 3)
        outputData(i, 4) = rawData(i, 4)
        outputData(i, 5) = rawData(i, 5)
        
        If entry.Validate() Then
            outputData(i, 6) = "Valid"
            validCount = validCount + 1
        Else
            outputData(i, 6) = "Invalid"
            invalidCount = invalidCount + 1
        End If
    Next i
    
    ' Clear old results, write headers, then write the WHOLE output array in one operation
    wsProcessed.Cells.Clear
    wsProcessed.Range("A1:F1").Value = Array("Employee Name", "Project Name", "Hours Logged", "Billable", "Work Date", "Status")
    wsProcessed.Range("A2").Resize(UBound(outputData, 1), 6).Value = outputData
    
    MsgBox "Processed " & (validCount + invalidCount) & " rows. Valid: " & validCount & ", Invalid: " & invalidCount
    Exit Sub

ErrHandler:
    Call LogError("ProcessTimesheets", Err.Number, Err.Description)
    MsgBox "An error occurred and was logged. Error: " & Err.Description
End Sub

Sub LogError(procName As String, errNum As Long, errDesc As String)
    Dim wsLog As Worksheet
    Dim lastRow As Long
    
    On Error Resume Next
    Set wsLog = ThisWorkbook.Sheets("ErrorLog")
    On Error GoTo 0
    
    If wsLog Is Nothing Then
        Set wsLog = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        wsLog.Name = "ErrorLog"
        wsLog.Range("A1:D1").Value = Array("Timestamp", "Procedure", "Error Number", "Error Description")
    End If
    
    lastRow = wsLog.Cells(wsLog.Rows.Count, 1).End(xlUp).Row + 1
    wsLog.Cells(lastRow, 1).Value = Now
    wsLog.Cells(lastRow, 2).Value = procName
    wsLog.Cells(lastRow, 3).Value = errNum
    wsLog.Cells(lastRow, 4).Value = errDesc
End Sub

Function UtilizationRate(billableHours As Double, totalHours As Double) As Variant
    If totalHours = 0 Then
        UtilizationRate = "N/A"
    Else
        UtilizationRate = Round(billableHours / totalHours * 100, 1)
    End If
End Function

Sub GenerateSummaryReport()
    On Error GoTo ErrHandler
    
    Dim wsProcessed As Worksheet
    Dim wsSummary As Worksheet
    Dim dict As New Dictionary
    Dim lastRow As Long
    Dim i As Long
    Dim empName As String
    Dim hrs As Double
    Dim key As Variant
    Dim outputRow As Long
    Dim cht As ChartObject
    
    Set wsProcessed = ThisWorkbook.Sheets("Processed")
    Set wsSummary = ThisWorkbook.Sheets("Summary")
    
    lastRow = wsProcessed.Cells(wsProcessed.Rows.Count, 1).End(xlUp).Row
    
    ' Aggregate billable hours per employee using a Dictionary
    For i = 2 To lastRow
        If wsProcessed.Cells(i, 6).Value = "Valid" And wsProcessed.Cells(i, 4).Value = "Y" Then
            empName = wsProcessed.Cells(i, 1).Value
            hrs = wsProcessed.Cells(i, 3).Value
            If dict.Exists(empName) Then
                dict(empName) = dict(empName) + hrs
            Else
                dict.Add empName, hrs
            End If
        End If
    Next i
    
    ' Clear old summary content and any old chart
    wsSummary.Cells.Clear
    For Each cht In wsSummary.ChartObjects
        cht.Delete
    Next cht
    
    ' Write the aggregated table
    wsSummary.Range("A1").Value = "Employee"
    wsSummary.Range("B1").Value = "Billable Hours"
    outputRow = 2
    For Each key In dict.Keys
        wsSummary.Cells(outputRow, 1).Value = key
        wsSummary.Cells(outputRow, 2).Value = dict(key)
        outputRow = outputRow + 1
    Next key
    
    ' Build a bar chart from that table
    Set cht = wsSummary.ChartObjects.Add(Left:=250, Width:=400, Top:=20, Height:=250)
    With cht.Chart
        .SetSourceData Source:=wsSummary.Range("A1:B" & (outputRow - 1))
        .ChartType = xlColumnClustered
        .HasTitle = True
        .ChartTitle.Text = "Billable Hours by Employee"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Employee"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "Billable Hours"
    End With
    
    Exit Sub
ErrHandler:
    Call LogError("GenerateSummaryReport", Err.Number, Err.Description)
    MsgBox "An error occurred and was logged. Error: " & Err.Description
End Sub
