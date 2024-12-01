Attribute VB_Name = "Module1"
Sub CreateMonthlySheets()
    ' Variables
    Dim sessionYear As String
    Dim month As Integer
    Dim ws As Worksheet
    Dim newSheetName As String
    Dim templateSheet As Worksheet
    Dim i As Integer
    
    ' Reference the template sheet
    Set templateSheet = ThisWorkbook.Sheets("Students Information")
    
    ' Use UserForm to select the session year
    UserForm1.Show
    sessionYear = UserForm1.SelectedSession
    If sessionYear = "" Then
        MsgBox "No session year selected. Aborting operation.", vbExclamation
        Exit Sub
    End If
    
    'Create Sheets
    For i = 1 To 12
        month = i
        newSheetName = monthName(month) & " " & sessionYear
        
        ' Clear the ws object at the start of each iteration
        Set ws = Nothing
        
        ' Check if the sheet already exists
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(newSheetName)
        On Error GoTo 0
        
        ' If the sheet exists, create it and apply the format and formulas
        If ws Is Nothing Then
            templateSheet.Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count)
            Set ws = ActiveSheet
            ws.Name = newSheetName
        Else
            Dim userResponse As VbMsgBoxResult
            userResponse = MsgBox("The sheet '" & newSheetName & "' already exists. Clearing its contents will remove all data and formatting. Do you want to continue?", vbYesNo + vbExclamation, "Warning")
            
            If userResponse = vbNo Then
                GoTo NextIteration ' Skip the rest of the code and move to the next iteration
            End If
            ws.Cells.Clear ' Clears all contents (values, formulas) and formatting
            templateSheet.Cells.Copy Destination:=ws.Cells(1, 1) ' Copy all cells starting from A1 in ws
        End If
        'Complete the template
        Call ApplyTemplateDesign(ws, month, sessionYear)
NextIteration:
    Next i
    
    MsgBox "Monthly sheets created for session: " & sessionYear, vbInformation
End Sub

Sub ApplyTemplateDesign(ws As Worksheet, month As Integer, sessionYear As String)
    Dim startDate As Date
    Dim daysInMonth As Integer
    Dim currentDay As Date
    Dim mName As String: mName = monthName(month)
    Dim year As Integer: year = SessionToYear(sessionYear, month)
    Dim i As Integer
    'Calculate lastRow
    Dim lastRow As Integer: lastRow = 2
    Dim myCell As Range
    'determine last row of current region
    Do While (Not IsEmpty(Cells(lastRow + 1, 1)))
        lastRow = lastRow + 1
    Loop
    
    ' Calculate the start date (1st of the month and session year)
    startDate = DateSerial(year, month, 1)
    daysInMonth = Day(DateSerial(year, month + 1, 0)) ' Get last day of the month
    
    
    ' Format of Date columns:
    
    ' Loop to add the dates (1st, 2nd, ..., last day of the month)
    For i = 1 To daysInMonth
        currentDay = startDate + (i - 1)
        With ws.Cells(2, i + 6)
            .Value = currentDay ' Add the date in row 1, starting from column C
            .NumberFormat = "dd" ' Format the date
            .ColumnWidth = 3.11
        End With
    Next i
    
    With ws.Range(ws.Cells(1, 7), ws.Cells(1, 6 + daysInMonth))
        .Merge
        .Value = "Attendance for " & monthName(month) & "-" & year
    End With
    
    With Range(ws.Cells(3, 7), ws.Cells(lastRow, 6 + daysInMonth))
            .Font.Name = "Perpetua Titling MT"
            .FormatConditions.Add(xlCellValue, xlEqual, "P").Interior.Color = 13561798
            .FormatConditions.Add(xlCellValue, xlEqual, "A").Interior.Color = 13551615
            .Borders.Weight = xlThin
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlEdgeBottom).Weight = xlMedium
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
    End With
    
    
    'Format of Total columns:
    
    With ws.Range(ws.Cells(1, 8 + daysInMonth), ws.Cells(1, 9 + daysInMonth))
        .Merge
        .Value = "Total"
        ColumnWidth = 5.89
    End With
    
    ws.Cells(2, 8 + daysInMonth).Value = "P"
    
    ws.Cells(2, 9 + daysInMonth).Value = "A"
    
    With Range(ws.Cells(3, 8 + daysInMonth), ws.Cells(lastRow, 9 + daysInMonth))
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
    End With
    
    For Each myCell In Range(ws.Cells(3, 8 + daysInMonth), ws.Cells(lastRow, 8 + daysInMonth))
        With myCell
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlEdgeBottom).Weight = xlThin
            .FormatConditions.Add(xlCellValue, xlGreater, 0).Interior.Color = 13561798
            .FormulaR1C1 = "=countifs(R" & myCell.Row & "C7:R" & myCell.Row & "C" & myCell.Column - 2 & ",""P"")"
        End With
    Next myCell
    
    For Each myCell In Range(ws.Cells(3, 9 + daysInMonth), ws.Cells(lastRow, 9 + daysInMonth))
        With myCell
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlEdgeBottom).Weight = xlThin
            .FormatConditions.Add(xlCellValue, xlGreater, 0).Interior.Color = 13551615
            .FormulaR1C1 = "=countifs(R" & myCell.Row & "C7:R" & myCell.Row & "C" & myCell.Column - 3 & ",""A"")"
        End With
    Next myCell
    
    With ws.Cells(lastRow + 1, 7 + daysInMonth)
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Value = "Total"
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217) ' Light blue background
    End With
    
    With Range(ws.Cells(lastRow + 1, 8 + daysInMonth), ws.Cells(lastRow + 1, 9 + daysInMonth))
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Value = "Grand Total"
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217) ' Light blue background
    End With
    
    With Range(ws.Cells(lastRow + 2, 8 + daysInMonth), ws.Cells(lastRow + 4, 8 + daysInMonth))
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .FormatConditions.Add(xlCellValue, xlGreater, 0).Interior.Color = 13561798
        .FormulaR1C1 = "=sum(R3C" & 8 + daysInMonth & ":R" & lastRow & "C" & 8 + daysInMonth & ")"
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    With Range(ws.Cells(lastRow + 2, 9 + daysInMonth), ws.Cells(lastRow + 4, 9 + daysInMonth))
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .FormatConditions.Add(xlCellValue, xlGreater, 0).Interior.Color = 13551615
        .FormulaR1C1 = "=sum(R3C" & 9 + daysInMonth & ":R" & lastRow & "C" & 9 + daysInMonth & ")"
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    
    With Range(ws.Cells(lastRow + 2, 1), ws.Cells(lastRow + 4, 4))
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeLeft).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Value = "Total"
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217) ' Light blue background
    End With
    
    With Range(ws.Cells(lastRow + 2, 5), ws.Cells(lastRow + 2, 6))
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeTop).Weight = xlMedium
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Value = "Female"
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217) ' Light blue background
    End With
    
    For Each myCell In Range(ws.Cells(lastRow + 2, 7), ws.Cells(lastRow + 2, 6 + daysInMonth))
        With myCell
            .FormatConditions.Add(xlCellValue, xlGreater, 0).Interior.Color = RGB(252, 228, 214)
            .FormulaR1C1 = "=countifs(R3C" & myCell.Column & ":R" & myCell.Row - 2 & "C" & myCell.Column & ",""P""," & "R3C5:R" & myCell.Row - 2 & "C5,""F"")"
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeTop).Weight = xlMedium
        End With
    Next myCell
    
    With Range(ws.Cells(lastRow + 3, 5), ws.Cells(lastRow + 3, 6))
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Value = "Male"
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217) ' Light blue background
    End With
    
    For Each myCell In Range(ws.Cells(lastRow + 3, 7), ws.Cells(lastRow + 3, 6 + daysInMonth))
        With myCell
            .FormatConditions.Add(xlCellValue, xlGreater, 0).Interior.Color = RGB(189, 215, 238)
            .FormulaR1C1 = "=countifs(R3C" & myCell.Column & ":R" & myCell.Row - 3 & "C" & myCell.Column & ",""P""," & "R3C5:R" & myCell.Row - 3 & "C5,""M"")"
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With
    Next myCell
    
    With Range(ws.Cells(lastRow + 4, 5), ws.Cells(lastRow + 4, 6))
        .Merge
        .HorizontalAlignment = xlCenterAcrossSelection
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeRight).Weight = xlMedium
        .Borders(xlEdgeBottom).Weight = xlMedium
        .Value = "Total"
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217) ' Light blue background
    End With
    
    For Each myCell In Range(ws.Cells(lastRow + 4, 7), ws.Cells(lastRow + 4, 6 + daysInMonth))
        With myCell
            .FormatConditions.Add(xlCellValue, xlGreater, 0).Interior.Color = 13561798
            .FormulaR1C1 = "=sum(R" & myCell.Row - 2 & "C" & myCell.Column & ":R" & myCell.Row - 1 & "C" & myCell.Column & ")"
            .Borders(xlEdgeRight).Weight = xlThin
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With
    Next myCell
    
    For Each myCell In Range(ws.Cells(lastRow + 2, 7 + daysInMonth), ws.Cells(lastRow + 4, 7 + daysInMonth))
        With myCell
            .FormatConditions.Add(xlCellValue, xlGreater, 0).Interior.Color = 13561798
            .FormulaR1C1 = "=sum(R" & myCell.Row & "C7:R" & myCell.Row & "C" & myCell.Column - 1 & ")"
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlEdgeBottom).Weight = xlMedium
        End With
    Next myCell
    
    'Format of Header Rows:
    
    With Range(ws.Cells(1, 1), ws.Cells(2, 9 + daysInMonth))
            .Font.Bold = True
            .Interior.Color = RGB(217, 217, 217) ' Light blue background
            .Borders.Weight = xlMedium
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
    End With
    
End Sub

' Helper function to convert sessionYear to year
Function SessionToYear(sessionYear As String, month As Integer) As Integer
    Dim years() As String
    years = Split(sessionYear, "-")
    If month < 4 Then
        SessionToYear = CInt(years(1))
    Else
        SessionToYear = CInt(years(0))
    End If
End Function
