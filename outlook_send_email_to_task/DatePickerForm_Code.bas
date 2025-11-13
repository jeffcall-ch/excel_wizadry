Option Explicit

' Form-level variables
Public SelectedDate As Date
Public Cancelled As Boolean
Public Priority As Integer

Private Sub UserForm_Initialize()
    Dim i As Integer
    
    ' Set default values
    Cancelled = True
    SelectedDate = Date
    Priority = 1 ' Default to Normal priority
    
    ' Set toggle button default state (False = Normal, True = High)
    tglPriority.Value = False
    tglPriority.Caption = "Normal Priority"
    
    ' Populate Month ComboBox
    cboMonth.Clear
    cboMonth.AddItem "January"
    cboMonth.AddItem "February"
    cboMonth.AddItem "March"
    cboMonth.AddItem "April"
    cboMonth.AddItem "May"
    cboMonth.AddItem "June"
    cboMonth.AddItem "July"
    cboMonth.AddItem "August"
    cboMonth.AddItem "September"
    cboMonth.AddItem "October"
    cboMonth.AddItem "November"
    cboMonth.AddItem "December"
    
    ' Populate Day ComboBox (1-31)
    cboDay.Clear
    For i = 1 To 31
        cboDay.AddItem i
    Next i
    
    ' Populate Year ComboBox (current year - 1 to current year + 5)
    cboYear.Clear
    For i = Year(Date) - 1 To Year(Date) + 5
        cboYear.AddItem i
    Next i
    
    ' Set default to today's date
    cboMonth.ListIndex = Month(Date) - 1
    cboDay.Value = Day(Date)
    cboYear.Value = Year(Date)
    
    ' Update label
    UpdateDateLabel
End Sub

Private Sub cboMonth_Change()
    UpdateDateLabel
    AdjustDaysInMonth
End Sub

Private Sub cboDay_Change()
    UpdateDateLabel
End Sub

Private Sub cboYear_Change()
    UpdateDateLabel
    AdjustDaysInMonth
End Sub

Private Sub AdjustDaysInMonth()
    ' Adjust days based on selected month/year
    Dim maxDays As Integer
    Dim currentDay As Integer
    
    If cboMonth.ListIndex = -1 Or cboYear.Value = "" Then Exit Sub
    
    currentDay = Val(cboDay.Value)
    maxDays = Day(DateSerial(Val(cboYear.Value), cboMonth.ListIndex + 2, 0))
    
    ' If current day exceeds max days in month, adjust it
    If currentDay > maxDays Then
        cboDay.Value = maxDays
    End If
End Sub

Private Sub UpdateDateLabel()
    Dim previewDate As String
    
    If cboMonth.ListIndex >= 0 And cboDay.Value <> "" And cboYear.Value <> "" Then
        previewDate = cboMonth.Value & " " & cboDay.Value & ", " & cboYear.Value
        lblPreview.Caption = "Selected: " & previewDate
    Else
        lblPreview.Caption = "Selected: (incomplete)"
    End If
End Sub

Private Sub tglPriority_Click()
    ' Toggle between Normal and High priority
    If tglPriority.Value Then
        tglPriority.Caption = "High Priority (!)"
        Priority = 2
    Else
        tglPriority.Caption = "Normal Priority"
        Priority = 1
    End If
End Sub

Private Sub btnOK_Click()
    ' Validate inputs
    If cboMonth.ListIndex = -1 Then
        MsgBox "Please select a month.", vbExclamation
        Exit Sub
    End If
    
    If cboDay.Value = "" Then
        MsgBox "Please select a day.", vbExclamation
        Exit Sub
    End If
    
    If cboYear.Value = "" Then
        MsgBox "Please select a year.", vbExclamation
        Exit Sub
    End If
    
    ' Create the date
    On Error GoTo InvalidDate
    SelectedDate = DateSerial(Val(cboYear.Value), cboMonth.ListIndex + 1, Val(cboDay.Value))
    Cancelled = False
    Me.Hide
    Exit Sub
    
InvalidDate:
    MsgBox "Invalid date selected. Please check your selection.", vbCritical
End Sub

Private Sub btnCancel_Click()
    Cancelled = True
    Me.Hide
End Sub

Private Sub btnToday_Click()
    ' Set to today's date
    cboMonth.ListIndex = Month(Date) - 1
    cboDay.Value = Day(Date)
    cboYear.Value = Year(Date)
End Sub
