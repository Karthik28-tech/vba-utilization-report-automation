VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmTimesheetEntry 
   Caption         =   "UserForm1"
   ClientHeight    =   3036
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   4584
   OleObjectBlob   =   "frmTimesheetEntry.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmTimesheetEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnSubmit_Click()
    Dim entry As New TimesheetEntry
    Dim wsProcessed As Worksheet
    Dim lastRow As Long
    Dim dict As New Dictionary
    Dim key As String
    Dim i As Long
    
    Set wsProcessed = ThisWorkbook.Sheets("Processed")
    
    ' Build a Dictionary of existing employee+project combos to check duplicates
    lastRow = wsProcessed.Cells(wsProcessed.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lastRow
        key = wsProcessed.Cells(i, 1).Value & "|" & wsProcessed.Cells(i, 2).Value & "|" & Format(wsProcessed.Cells(i, 5).Value, "yyyy-mm-dd")
        If Not dict.Exists(key) Then
            dict.Add key, True
        End If
    Next i
    
    ' Load form values into the class
    entry.EmployeeName = Me.txtEmployeeName.Value
    entry.ProjectName = Me.txtProjectName.Value
    If IsNumeric(Me.txtHours.Value) Then entry.HoursLogged = CDbl(Me.txtHours.Value)
    entry.Billable = UCase(Me.txtBillable.Value)
    If IsDate(Me.txtWorkDate.Value) Then entry.WorkDate = CDate(Me.txtWorkDate.Value)
    
    ' Check for duplicate before validating
    key = entry.EmployeeName & "|" & entry.ProjectName & "|" & Format(entry.WorkDate, "yyyy-mm-dd")
    If dict.Exists(key) Then
        MsgBox "This employee already has a timesheet entry for this project on this date."
        Exit Sub
    End If
    
    ' Validate
    If Not entry.Validate() Then
        MsgBox "Entry is invalid. Please check all fields (hours must be 0-24, billable must be Y or N, all fields required)."
        Exit Sub
    End If
    
    ' Write to Processed sheet
    lastRow = wsProcessed.Cells(wsProcessed.Rows.Count, 1).End(xlUp).Row + 1
    wsProcessed.Cells(lastRow, 1).Value = entry.EmployeeName
    wsProcessed.Cells(lastRow, 2).Value = entry.ProjectName
    wsProcessed.Cells(lastRow, 3).Value = entry.HoursLogged
    wsProcessed.Cells(lastRow, 4).Value = entry.Billable
    wsProcessed.Cells(lastRow, 5).Value = entry.WorkDate
    wsProcessed.Cells(lastRow, 6).Value = "Valid"
    
    MsgBox "Entry added successfully."
    
    ' Clear the form
    Me.txtEmployeeName.Value = ""
    Me.txtProjectName.Value = ""
    Me.txtHours.Value = ""
    Me.txtBillable.Value = ""
    Me.txtWorkDate.Value = ""
End Sub

Private Sub Label4_Click()

End Sub
