VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TimesheetEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public EmployeeName As String
Public ProjectName As String
Public HoursLogged As Double
Public Billable As String
Public WorkDate As Date

Public Function Validate() As Boolean
    Dim isValid As Boolean
    isValid = True
    
    If EmployeeName = "" Then
        isValid = False
    End If
    
    If HoursLogged <= 0 Or HoursLogged > 24 Then
        isValid = False
    End If
    
    If Billable <> "Y" And Billable <> "N" Then
        isValid = False
    End If
    
    If WorkDate = 0 Then
        isValid = False
    End If
    
    Validate = isValid
End Function
