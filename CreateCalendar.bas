Sub CreateCalendar()
    Dim yearInput As String
    Dim calYear As Integer
    Dim ws As Worksheet
    Dim month As Integer
    Dim dayNum As Integer
    Dim firstDay As Date
    Dim lastDay As Date
    Dim row As Integer, col As Integer
    Dim dayNames(0 To 6) As String
    Dim monthNames(1 To 12) As String
    Dim holidays As Collection
    Dim workingDaysByMonth(1 To 12) As Integer
    Dim totalHolidays As Integer
    Dim totalVacationDays As Integer
    Dim totalWorkingDays As Integer
    
    ' Day names in English, starting with Monday
    dayNames(0) = "Monday"
    dayNames(1) = "Tuesday"
    dayNames(2) = "Wednesday"
    dayNames(3) = "Thursday"
    dayNames(4) = "Friday"
    dayNames(5) = "Saturday"
    dayNames(6) = "Sunday"
    
    ' Month names in English
    monthNames(1) = "January"
    monthNames(2) = "February"
    monthNames(3) = "March"
    monthNames(4) = "April"
    monthNames(5) = "May"
    monthNames(6) = "June"
    monthNames(7) = "July"
    monthNames(8) = "August"
    monthNames(9) = "September"
    monthNames(10) = "October"
    monthNames(11) = "November"
    monthNames(12) = "December"
    
    ' Read year from merged cell C2
    yearInput = Range("C2").Value
    
    If yearInput = "" Then
        MsgBox "Please enter a year in cell C2.", vbCritical
        Exit Sub
    End If
    
    ' Validate year input
    On Error Resume Next
    calYear = CInt(yearInput)
    On Error GoTo 0
    
    If calYear < 1900 Or calYear > 9999 Then
        MsgBox "Please enter a valid year between 1900 and 9999.", vbCritical
        Exit Sub
    End If
    
    ' Get French holidays
    Set holidays = GetFrenchHolidays(calYear)
    
    ' Define active worksheet
    Set ws = ActiveSheet
    
    ' Enable events
    ws.Name = "CalendarSheet" + "_" + yearInput
    Application.EnableEvents = True
    
    ws.Range("B:AH").ColumnWidth = 12
    
    ' The calendars are displayed horizontally (3 per row) starting from B23
    Dim startRow As Integer
    Dim startCol As Integer
    
    For month = 1 To 12
        ' Calculate row and column for horizontal layout (3 calendars per row)
        Dim monthInRow As Integer
        Dim monthRowNumber As Integer
        monthInRow = ((month - 1) Mod 3)      ' 0, 1, or 2: position in row (3 calendars per row)
        monthRowNumber = (month - 1) \ 3      ' Which row of calendars
        startRow = 26 + (monthRowNumber * 13)  ' Start row for this set of months (13 rows = calendar + 1 blank line)
        startCol = 2 + (monthInRow * 8)        ' Start column for this calendar (8 = 7 columns + 1 separator)
        ' Month title - merge B to H
        ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 6)).Merge
        ws.Cells(startRow, startCol).Value = monthNames(month) & " " & calYear
        ws.Cells(startRow, startCol).Font.Bold = True
        ws.Cells(startRow, startCol).Font.Size = 14
        ws.Cells(startRow, startCol).HorizontalAlignment = xlCenter
        
        ' Border around title
        With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 6)).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
        startRow = startRow + 1
        
        ' Day headers
        Dim i As Integer
        For i = 0 To 6
            ws.Cells(startRow, startCol + i).Value = dayNames(i)
            ws.Cells(startRow, startCol + i).Font.Bold = True
            ws.Cells(startRow, startCol + i).HorizontalAlignment = xlCenter
            ' Blue color for day headers
            ws.Cells(startRow, startCol + i).Interior.Color = RGB(91, 155, 213)
            ws.Cells(startRow, startCol + i).Font.Color = RGB(255, 255, 255)  ' White text
            
            ' Border around each header
            With ws.Cells(startRow, startCol + i).Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        Next i
        startRow = startRow + 1
        
        ' Remember start of dates for this month
        Dim dateStartRow As Integer
        dateStartRow = startRow
        
        ' Get first and last day of month
        firstDay = DateSerial(calYear, month, 1)
        lastDay = DateSerial(calYear, month + 1, 0)
        
        ' Get day of week for first day (1=Sunday, 7=Saturday)
        ' Convert to start on Monday
        Dim weekStartCol As Integer
        weekStartCol = Weekday(firstDay) - 2
        If weekStartCol = -1 Then
            weekStartCol = 6
        End If
        
        ' Fill in dates
        col = weekStartCol
        row = startRow
        
        For dayNum = 1 To day(lastDay)
            Dim cellDate As Date
            cellDate = DateSerial(calYear, month, dayNum)
            
            ws.Cells(row, startCol + col).Value = dayNum
            ws.Cells(row, startCol + col).HorizontalAlignment = xlRight
            ws.Cells(row, startCol + col).VerticalAlignment = xlTop
            ws.Cells(row, startCol + col).RowHeight = 15
            
            ' Check if it's a holiday
            Dim isHoliday As Boolean
            Dim j As Integer
            Dim holidayName As String
            isHoliday = False
            For j = 1 To holidays.Count
                If holidays(j) = cellDate Then
                    isHoliday = True
                    Exit For
                End If
            Next j
            
            ' Color differently holidays and weekends
            ' Check if cell is not already green
            Dim isGreen As Boolean
            isGreen = False
            If ws.Cells(row, startCol + col).Interior.Color = RGB(0, 176, 0) Or ws.Cells(row, startCol + col).Interior.Color = RGB(0, 176, 80) Or _
               ws.Cells(row, startCol + col).Interior.Color = RGB(146, 208, 80) Or ws.Cells(row, startCol + col).Interior.Color = RGB(0, 255, 0) Then
                isGreen = True
            End If
            
            If Not isGreen Then
                If isHoliday Then
                    ws.Cells(row, startCol + col).Interior.Color = RGB(255, 200, 0)  ' Orange for holidays
                    ws.Cells(row, startCol + col).Font.Bold = True
                    
                    ' Add a note with the holiday name
                    holidayName = GetHolidayName(cellDate, calYear)
                    On Error Resume Next
                    ws.Cells(row, startCol + col).AddComment holidayName
                    On Error GoTo 0
                ElseIf col = 5 Or col = 6 Then
                    ws.Cells(row, startCol + col).Interior.Color = RGB(220, 220, 220)
                End If
            End If
            
            ' Border around each date cell
            With ws.Cells(row, startCol + col).Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            
            col = col + 1
            If col > 6 Then
                col = 0
                row = row + 1
            End If
        Next dayNum
        
        ' Add working and vacation days count below the month
        Dim workingDaysCount As Integer
        Dim vacationDaysCount As Integer
        Dim summaryCol As Integer
        Dim summaryRow As Integer
        Dim dateRangeForMonth As Range
        summaryCol = startCol + 1  ' Starting from column C (column 3)
        summaryRow = row + 2       ' Below the calendar dates with a blank line
        
        ' Count vacation days (red cells)
        Set dateRangeForMonth = ws.Range(ws.Cells(dateStartRow, startCol), ws.Cells(row, startCol + 6))
        vacationDaysCount = CountVacationDays(dateRangeForMonth)
        totalVacationDays = totalVacationDays + vacationDaysCount
        
        ' Calculate working days by subtracting vacation days and adding green days
        Dim greenDaysCount As Integer
        greenDaysCount = CountGreenDays(dateRangeForMonth)
        workingDaysCount = CountWorkingDays(month, calYear, holidays) - vacationDaysCount + greenDaysCount
        totalWorkingDays = totalWorkingDays + workingDaysCount
        workingDaysByMonth(month) = workingDaysCount
        ws.Cells(summaryRow, summaryCol).Value = "Working days:"
        ws.Cells(summaryRow, summaryCol).Font.Bold = True
        ws.Cells(summaryRow, summaryCol + 1).Value = workingDaysCount
        ws.Cells(summaryRow, summaryCol + 1).Font.Bold = True
        summaryRow = summaryRow + 1
        
        ws.Cells(summaryRow, summaryCol).Value = "Days off:"
        ws.Cells(summaryRow, summaryCol).Font.Bold = True
        ws.Cells(summaryRow, summaryCol + 1).Value = vacationDaysCount
        ws.Cells(summaryRow, summaryCol + 1).Font.Bold = True
        summaryRow = summaryRow + 1
        
        ' Count holidays in the month
        Dim holidaysCount As Integer
        holidaysCount = CountHolidaysInMonth(holidays, month, calYear)
        totalHolidays = totalHolidays + holidaysCount
        ws.Cells(summaryRow, summaryCol).Value = "Holidays:"
        ws.Cells(summaryRow, summaryCol).Font.Bold = True
        ws.Cells(summaryRow, summaryCol + 1).Value = holidaysCount
        ws.Cells(summaryRow, summaryCol + 1).Font.Bold = True
        
        ' Add borders to summary cells
        Dim summaryBorderRow As Integer
        For summaryBorderRow = row + 2 To summaryRow
            With ws.Cells(summaryBorderRow, summaryCol).Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
            With ws.Cells(summaryBorderRow, summaryCol + 1).Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        Next summaryBorderRow
        
        ' Column width for working and vacation days
        ws.Columns(summaryCol).ColumnWidth = 15
    Next month
    
    ' Add total holidays in L2
    ws.Cells(1, 12).Value = "Total Holidays:"
    ws.Cells(1, 12).Font.Bold = True
    ws.Cells(2, 12).Value = totalHolidays
    ws.Cells(2, 12).Font.Bold = True
    ws.Cells(2, 12).NumberFormat = "0"
    
    ' Add total vacation days in L3
    ws.Cells(3, 12).Value = totalVacationDays
    ws.Cells(3, 12).Font.Bold = True
    ws.Cells(3, 12).NumberFormat = "0"
    
    ' Add total working days in L4
    ws.Cells(4, 12).Value = totalWorkingDays
    ws.Cells(4, 12).Font.Bold = True
    ws.Cells(4, 12).NumberFormat = "0"
    
    ' Create summary table starting from F7
    Dim tableCol As Integer
    Dim tableRow As Integer
    Dim multiplierValue As Double
    tableCol = 6   ' Column F
    tableRow = 7   ' Starting row 8
    
    ' Merge cells F6:F7, G6:G7, etc. for headers
    Dim mergedRange As Range
    
    Set mergedRange = ws.Range(ws.Cells(6, tableCol), ws.Cells(7, tableCol))
    mergedRange.Merge
    mergedRange.Value = "Month"
    mergedRange.Font.Bold = True
    mergedRange.Interior.Color = RGB(91, 155, 213)
    mergedRange.Font.Color = RGB(255, 255, 255)
    mergedRange.HorizontalAlignment = xlCenter
    
    Set mergedRange = ws.Range(ws.Cells(6, tableCol + 1), ws.Cells(7, tableCol + 1))
    mergedRange.Merge
    mergedRange.Value = "Working days"
    mergedRange.Font.Bold = True
    mergedRange.Interior.Color = RGB(91, 155, 213)
    mergedRange.Font.Color = RGB(255, 255, 255)
    mergedRange.HorizontalAlignment = xlCenter
    
    Set mergedRange = ws.Range(ws.Cells(6, tableCol + 2), ws.Cells(7, tableCol + 2))
    mergedRange.Merge
    mergedRange.Value = "Revenue"
    mergedRange.Font.Bold = True
    mergedRange.Interior.Color = RGB(91, 155, 213)
    mergedRange.Font.Color = RGB(255, 255, 255)
    mergedRange.HorizontalAlignment = xlCenter
    
    Set mergedRange = ws.Range(ws.Cells(6, tableCol + 3), ws.Cells(7, tableCol + 3))
    mergedRange.Merge
    mergedRange.Value = "Monthly Load"
    mergedRange.Font.Bold = True
    mergedRange.Interior.Color = RGB(91, 155, 213)
    mergedRange.Font.Color = RGB(255, 255, 255)
    mergedRange.HorizontalAlignment = xlCenter
    
    Set mergedRange = ws.Range(ws.Cells(6, tableCol + 4), ws.Cells(7, tableCol + 4))
    mergedRange.Merge
    mergedRange.Value = "MAY"
    mergedRange.Font.Bold = True
    mergedRange.Interior.Color = RGB(91, 155, 213)
    mergedRange.Font.Color = RGB(255, 255, 255)
    mergedRange.HorizontalAlignment = xlCenter
    
    Set mergedRange = ws.Range(ws.Cells(6, tableCol + 5), ws.Cells(7, tableCol + 5))
    mergedRange.Merge
    mergedRange.Value = "Available Balance"
    mergedRange.Font.Bold = True
    mergedRange.Interior.Color = RGB(91, 155, 213)
    mergedRange.Font.Color = RGB(255, 255, 255)
    mergedRange.HorizontalAlignment = xlCenter
    
    Set mergedRange = ws.Range(ws.Cells(6, tableCol + 6), ws.Cells(7, tableCol + 6))
    mergedRange.Merge
    mergedRange.Value = "Professional expense"
    mergedRange.Font.Bold = True
    mergedRange.Interior.Color = RGB(91, 155, 213)
    mergedRange.Font.Color = RGB(255, 255, 255)
    mergedRange.HorizontalAlignment = xlCenter
    
    Set mergedRange = ws.Range(ws.Cells(6, tableCol + 7), ws.Cells(7, tableCol + 7))
    mergedRange.Merge
    mergedRange.Value = "Treasury"
    mergedRange.Font.Bold = True
    mergedRange.Interior.Color = RGB(91, 155, 213)
    mergedRange.Font.Color = RGB(255, 255, 255)
    mergedRange.HorizontalAlignment = xlCenter
    
    Set mergedRange = ws.Range(ws.Cells(6, tableCol + 8), ws.Cells(7, tableCol + 8))
    mergedRange.Merge
    mergedRange.Value = "Net Salary"
    mergedRange.Font.Bold = True
    mergedRange.Interior.Color = RGB(198, 239, 206)
    mergedRange.Font.Color = RGB(0, 97, 0)
    mergedRange.HorizontalAlignment = xlCenter
    
    Set mergedRange = ws.Range(ws.Cells(6, tableCol + 9), ws.Cells(7, tableCol + 9))
    mergedRange.Merge
    mergedRange.Value = "Meal vouchers"
    mergedRange.Font.Bold = True
    mergedRange.Interior.Color = RGB(198, 239, 206)
    mergedRange.Font.Color = RGB(0, 97, 0)
    mergedRange.HorizontalAlignment = xlCenter
    
    ' Get value from C7
    multiplierValue = ws.Range("C7").Value
    
    ' Add borders to table headers
    With ws.Cells(tableRow, tableCol).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Cells(tableRow, tableCol + 1).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Cells(tableRow, tableCol + 2).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Cells(tableRow, tableCol + 3).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Cells(tableRow, tableCol + 4).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Cells(tableRow, tableCol + 5).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Cells(tableRow, tableCol + 6).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Cells(tableRow, tableCol + 7).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Cells(tableRow, tableCol + 8).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With ws.Cells(tableRow, tableCol + 9).Borders
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    ' Fill table data
    tableRow = tableRow + 1
    For month = 1 To 12
        ws.Cells(tableRow, tableCol).Value = monthNames(month)
        ws.Cells(tableRow, tableCol + 1).Value = workingDaysByMonth(month)
        ws.Cells(tableRow, tableCol + 2).Formula = "=$C$5*" & ws.Cells(tableRow, tableCol + 1).Address
        ws.Cells(tableRow, tableCol + 2).NumberFormat = "#,##0.00 ""€"""
        ws.Cells(tableRow, tableCol + 3).Formula = "=" & ws.Cells(tableRow, tableCol + 2).Address & "*($C$7+$C$8)+IF(" & ws.Cells(tableRow, tableCol + 4).Address & "<>0,$C$9,0)+$C$12"
        ws.Cells(tableRow, tableCol + 3).NumberFormat = "#,##0.00 ""€"""
        
        ' Keep existing MAY value if different from 1500, otherwise set to 1500
        If ws.Cells(tableRow, tableCol + 4).Value <> "" And ws.Cells(tableRow, tableCol + 4).Value <> 1500 Then
            ' Keep the existing value
        Else
            ws.Cells(tableRow, tableCol + 4).Value = 1500
        End If
        ws.Cells(tableRow, tableCol + 4).NumberFormat = "#,##0.00 ""€"""
        ws.Cells(tableRow, tableCol + 5).Formula = "=" & ws.Cells(tableRow, tableCol + 2).Address & "-" & ws.Cells(tableRow, tableCol + 3).Address & "-" & ws.Cells(tableRow, tableCol + 4).Address
        ws.Cells(tableRow, tableCol + 5).NumberFormat = "#,##0.00 ""€"""
        If ws.Cells(tableRow, tableCol + 6).Value = "" Then
            ws.Cells(tableRow, tableCol + 6).Value = 0
        End If
        ws.Cells(tableRow, tableCol + 6).NumberFormat = "#,##0.00 ""€"""
        ws.Cells(tableRow, tableCol + 7).Formula = "=" & ws.Cells(tableRow, tableCol + 5).Address & "-" & ws.Cells(tableRow, tableCol + 6).Address & "-$C$16-$C$20-$C$22"
        ws.Cells(tableRow, tableCol + 7).NumberFormat = "#,##0.00 ""€"";[Red]-#,##0.00 ""€"""
        ws.Cells(tableRow, tableCol + 8).Formula = "=$C$16-$C$21-($C$17*" & ws.Cells(tableRow, tableCol + 1).Address & ")/2-$C$12/2"
        ws.Cells(tableRow, tableCol + 8).NumberFormat = "#,##0.00 ""€"""
        ws.Cells(tableRow, tableCol + 8).Font.Color = RGB(250, 125, 0)
        ws.Cells(tableRow, tableCol + 8).Interior.Color = RGB(242, 242, 242)
        ws.Cells(tableRow, tableCol + 9).Formula = "=$C$17*" & ws.Cells(tableRow, tableCol + 1).Address
        ws.Cells(tableRow, tableCol + 9).NumberFormat = "#,##0.00 ""€"""
        ws.Cells(tableRow, tableCol + 9).Font.Color = RGB(250, 125, 0)
        ws.Cells(tableRow, tableCol + 9).Interior.Color = RGB(242, 242, 242)
        
        ' Add borders
        With ws.Cells(tableRow, tableCol).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With ws.Cells(tableRow, tableCol + 1).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With ws.Cells(tableRow, tableCol + 2).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With ws.Cells(tableRow, tableCol + 3).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With ws.Cells(tableRow, tableCol + 4).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With ws.Cells(tableRow, tableCol + 5).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With ws.Cells(tableRow, tableCol + 6).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With ws.Cells(tableRow, tableCol + 7).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With ws.Cells(tableRow, tableCol + 8).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        With ws.Cells(tableRow, tableCol + 9).Borders
            .LineStyle = xlContinuous
            .Weight = xlThin
        End With
        
        tableRow = tableRow + 1
    Next month
    
    ' Set column width for table
    ws.Columns(tableCol).ColumnWidth = 10
    ws.Columns(tableCol + 1).ColumnWidth = 10
    ws.Columns(tableCol + 2).ColumnWidth = 10
    ws.Columns(tableCol + 3).ColumnWidth = 12
    ws.Columns(tableCol + 4).ColumnWidth = 10
    ws.Columns(tableCol + 5).ColumnWidth = 14
    ws.Columns(tableCol + 6).ColumnWidth = 15
    ws.Columns(tableCol + 7).ColumnWidth = 12
    ws.Columns(tableCol + 8).ColumnWidth = 12
    ws.Columns(tableCol + 9).ColumnWidth = 12
    
    MsgBox "The calendar for " & calYear & " has been created!", vbInformation
End Sub

Function GetFrenchHolidays(year As Integer) As Collection
    Set GetFrenchHolidays = New Collection
    
    ' Fixed holidays
    GetFrenchHolidays.Add DateSerial(year, 1, 1)    ' New Year's Day
    GetFrenchHolidays.Add DateSerial(year, 5, 1)    ' Labour Day
    GetFrenchHolidays.Add DateSerial(year, 5, 8)    ' Victory Day 1945
    GetFrenchHolidays.Add DateSerial(year, 7, 14)   ' Bastille Day
    GetFrenchHolidays.Add DateSerial(year, 8, 15)   ' Assumption of Mary
    GetFrenchHolidays.Add DateSerial(year, 11, 1)   ' All Saints' Day
    GetFrenchHolidays.Add DateSerial(year, 11, 11)  ' Armistice Day
    GetFrenchHolidays.Add DateSerial(year, 12, 25)  ' Christmas
    
    ' Moveable holidays based on Easter
    Dim easterDate As Date
    easterDate = CalculateEaster(year)
    
    GetFrenchHolidays.Add easterDate + 1            ' Easter Monday
    GetFrenchHolidays.Add easterDate + 39           ' Ascension Day
    GetFrenchHolidays.Add easterDate + 50           ' Whit Monday
End Function

Function CalculateEaster(year As Integer) As Date
    ' Computus algorithm to calculate Easter date
    Dim a As Integer, b As Integer, c As Integer, d As Integer
    Dim e As Integer, f As Integer, g As Integer, h As Integer
    Dim i As Integer, k As Integer, l As Integer, m As Integer
    Dim month As Integer, day As Integer
    
    a = year Mod 19
    b = year \ 100
    c = year Mod 100
    d = b \ 4
    e = b Mod 4
    f = (b + 8) \ 25
    g = (b - f + 1) \ 3
    h = (19 * a + b - d - g + 15) Mod 30
    i = c \ 4
    k = c Mod 4
    l = (32 + 2 * e + 2 * i - h - k) Mod 7
    m = (a + 11 * h + 22 * l) \ 451
    month = (h + l - 7 * m + 114) \ 31
    day = ((h + l - 7 * m + 114) Mod 31) + 1
    
    CalculateEaster = DateSerial(year, month, day)
End Function

Function CountHolidaysInMonth(holidays As Collection, targetMonth As Integer, year As Integer) As Integer
    ' Count holidays in the month
    Dim count As Integer
    Dim j As Integer
    Dim holidayMonth As Integer
    count = 0
    
    For j = 1 To holidays.Count
        holidayMonth = Month(holidays(j))
        If holidayMonth = targetMonth Then
            count = count + 1
        End If
    Next j
    
    CountHolidaysInMonth = count
End Function

Function GetHolidayName(cellDate As Date, year As Integer) As String
    ' Return the name of the holiday for a given date
    Dim easterDate As Date
    
    ' Fixed holidays
    If cellDate = DateSerial(year, 1, 1) Then GetHolidayName = "New Year's Day": Exit Function
    If cellDate = DateSerial(year, 5, 1) Then GetHolidayName = "Labour Day": Exit Function
    If cellDate = DateSerial(year, 5, 8) Then GetHolidayName = "Victory Day 1945": Exit Function
    If cellDate = DateSerial(year, 7, 14) Then GetHolidayName = "Bastille Day": Exit Function
    If cellDate = DateSerial(year, 8, 15) Then GetHolidayName = "Assumption of Mary": Exit Function
    If cellDate = DateSerial(year, 11, 1) Then GetHolidayName = "All Saints' Day": Exit Function
    If cellDate = DateSerial(year, 11, 11) Then GetHolidayName = "Armistice Day": Exit Function
    If cellDate = DateSerial(year, 12, 25) Then GetHolidayName = "Christmas": Exit Function
    
    ' Moveable holidays based on Easter
    easterDate = CalculateEaster(year)
    If cellDate = easterDate + 1 Then GetHolidayName = "Easter Monday": Exit Function
    If cellDate = easterDate + 39 Then GetHolidayName = "Ascension Day": Exit Function
    If cellDate = easterDate + 50 Then GetHolidayName = "Whit Monday": Exit Function
    
    GetHolidayName = ""
End Function

Function CountWorkingDays(month As Integer, year As Integer, holidays As Collection) As Integer
    ' Count working days in the month (Monday to Friday, excluding holidays)
    Dim count As Integer
    Dim dayNum As Integer
    Dim lastDay As Integer
    Dim cellDate As Date
    Dim dayOfWeek As Integer
    Dim isHoliday As Boolean
    Dim j As Integer
    
    count = 0
    lastDay = day(DateSerial(year, month + 1, 0))
    
    For dayNum = 1 To lastDay
        cellDate = DateSerial(year, month, dayNum)
        dayOfWeek = Weekday(cellDate) - 2
        If dayOfWeek = -1 Then dayOfWeek = 6
        
        ' Check if it's a working day (Monday to Friday, col 0 to 4)
        If dayOfWeek >= 0 And dayOfWeek <= 4 Then
            ' Check if it's not a holiday
            isHoliday = False
            For j = 1 To holidays.Count
                If holidays(j) = cellDate Then
                    isHoliday = True
                    Exit For
                End If
            Next j
            
            If Not isHoliday Then
                count = count + 1
            End If
        End If
    Next dayNum
    
    CountWorkingDays = count
End Function

Function CountVacationDays(dateRange As Range) As Integer
    ' Count red cells (vacation days)
    Dim cell As Range
    Dim count As Integer
    count = 0
    
    On Error Resume Next
    For Each cell In dateRange
        If cell.Interior.Color = RGB(255, 0, 0) Or cell.Interior.Color = RGB(192, 0, 0) Or _
           cell.Interior.Color = RGB(255, 192, 203) Then
            count = count + 1
        End If
    Next cell
    On Error GoTo 0
    
    CountVacationDays = count
End Function

Function CountGreenDays(dateRange As Range) As Integer
    ' Count green cells (additional working days)
    Dim cell As Range
    Dim count As Integer
    count = 0
    
    On Error Resume Next
    For Each cell In dateRange
        If cell.Interior.Color = RGB(0, 176, 0) Or cell.Interior.Color = RGB(0, 176, 80) Or _
           cell.Interior.Color = RGB(146, 208, 80) Or cell.Interior.Color = RGB(0, 255, 0) Then
            count = count + 1
        End If
    Next cell
    On Error GoTo 0
    
    CountGreenDays = count
End Function

