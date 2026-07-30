VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    Dim wsProcessed As Worksheet
    Dim invalidCount As Long
    Dim i As Long
    Dim lastRow As Long
    
    Set wsProcessed = ThisWorkbook.Sheets("Processed")
    lastRow = wsProcessed.Cells(wsProcessed.Rows.Count, 1).End(xlUp).Row
    invalidCount = 0
    
    For i = 2 To lastRow
        If wsProcessed.Cells(i, 6).Value = "Invalid" Then
            invalidCount = invalidCount + 1
        End If
    Next i
    
    If invalidCount > 0 Then
        MsgBox invalidCount & " invalid timesheet entries need review.", vbExclamation
    End If
End Sub
