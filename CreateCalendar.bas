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
    
    ' Noms des jours en anglais, commençant par lundi
    dayNames(0) = "Monday"
    dayNames(1) = "Tuesday"
    dayNames(2) = "Wednesday"
    dayNames(3) = "Thursday"
    dayNames(4) = "Friday"
    dayNames(5) = "Saturday"
    dayNames(6) = "Sunday"
    
    ' Noms des mois en anglais
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
    
    ' Lire l'année à partir de la cellule fusionnée C2
    yearInput = Range("C3").Value
    
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
    
    ' Récupérer les jours fériés français
    Set holidays = GetFrenchHolidays(calYear)
    
    ' Définir le worksheet actif
    Set ws = ActiveSheet
    
    ' Activer les événements de changement
    ws.Name = "CalendarSheet"
    Application.EnableEvents = True
    
    ws.Range("A:G").ColumnWidth = 10
    
    ' Le calendrier commence à partir de B9
    Dim startRow As Integer
    startRow = 9
    Dim startCol As Integer
    startCol = 2  ' Colonne B
    
    For month = 1 To 12
        ' Month title - fusionner B à H
        ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 6)).Merge
        ws.Cells(startRow, startCol).Value = monthNames(month) & " " & calYear
        ws.Cells(startRow, startCol).Font.Bold = True
        ws.Cells(startRow, startCol).Font.Size = 14
        ws.Cells(startRow, startCol).HorizontalAlignment = xlCenter
        
        ' Bordure autour du titre
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
            ' Couleur bleu #5B9BD5 pour l'en-tête des jours
            ws.Cells(startRow, startCol + i).Interior.Color = RGB(91, 155, 213)
            ws.Cells(startRow, startCol + i).Font.Color = RGB(255, 255, 255)  ' Texte blanc
            
            ' Bordure autour de chaque en-tête
            With ws.Cells(startRow, startCol + i).Borders
                .LineStyle = xlContinuous
                .Weight = xlThin
            End With
        Next i
        startRow = startRow + 1
        
        ' Mémoriser le début des dates pour ce mois
        Dim dateStartRow As Integer
        dateStartRow = startRow
        
        ' Get first and last day of month
        firstDay = DateSerial(calYear, month, 1)
        lastDay = DateSerial(calYear, month + 1, 0)
        
        ' Get day of week for first day (1=Sunday, 7=Saturday)
        ' Convertir pour commencer au lundi
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
            
            ' Vérifier si c'est un jour férié
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
            
            ' Colorier différemment les jours fériés et weekends
            If isHoliday Then
                ws.Cells(row, startCol + col).Interior.Color = RGB(255, 200, 0)  ' Orange pour jours fériés
                ws.Cells(row, startCol + col).Font.Bold = True
                
                ' Ajouter une note avec le nom du jour férié
                holidayName = GetHolidayName(cellDate, calYear)
                On Error Resume Next
                ws.Cells(row, startCol + col).AddComment holidayName
                On Error GoTo 0
            ElseIf col = 5 Or col = 6 Then
                ws.Cells(row, startCol + col).Interior.Color = RGB(220, 220, 220)
            End If
            
            ' Bordure autour de chaque cellule de date
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
        
        ' Ajouter le nombre de jours travaillés et de jours en congé à droite du mois
        Dim workingDaysCount As Integer
        Dim vacationDaysCount As Integer
        Dim summaryCol As Integer
        Dim summaryRow As Integer
        Dim dateRangeForMonth As Range
        summaryCol = startCol + 8  ' Colonne I
        summaryRow = startRow
        
        ' Compter les jours en congé (cellules rouges)
        Set dateRangeForMonth = ws.Range(ws.Cells(dateStartRow, startCol), ws.Cells(row, startCol + 6))
        vacationDaysCount = CountVacationDays(dateRangeForMonth)
        
        ' Calculer les jours travaillés en soustrayant les jours de congé
        workingDaysCount = CountWorkingDays(month, calYear, holidays) - vacationDaysCount
        ws.Cells(summaryRow, summaryCol).Value = "Jours travaillés: " & workingDaysCount
        ws.Cells(summaryRow, summaryCol).Font.Bold = True
        summaryRow = summaryRow + 1
        
        ws.Cells(summaryRow, summaryCol).Value = "Jours en congé: " & vacationDaysCount
        ws.Cells(summaryRow, summaryCol).Font.Bold = True
        
        startRow = row + 3
    Next month
    
    MsgBox "Le calendrier pour " & calYear & " a été créé!", vbInformation
End Sub

Function GetFrenchHolidays(year As Integer) As Collection
    Set GetFrenchHolidays = New Collection
    
    ' Jours fériés fixes
    GetFrenchHolidays.Add DateSerial(year, 1, 1)    ' Jour de l'An
    GetFrenchHolidays.Add DateSerial(year, 5, 1)    ' Fête du Travail
    GetFrenchHolidays.Add DateSerial(year, 5, 8)    ' Fête de la Victoire 1945
    GetFrenchHolidays.Add DateSerial(year, 7, 14)   ' Fête nationale
    GetFrenchHolidays.Add DateSerial(year, 8, 15)   ' Assomption
    GetFrenchHolidays.Add DateSerial(year, 11, 1)   ' Toussaint
    GetFrenchHolidays.Add DateSerial(year, 11, 11)  ' Armistice
    GetFrenchHolidays.Add DateSerial(year, 12, 25)  ' Noël
    
    ' Jours fériés mobiles basés sur Pâques
    Dim easterDate As Date
    easterDate = CalculateEaster(year)
    
    GetFrenchHolidays.Add easterDate + 1            ' Lundi de Pâques
    GetFrenchHolidays.Add easterDate + 39           ' Ascension
    GetFrenchHolidays.Add easterDate + 50           ' Lundi de Pentecôte
End Function

Function CalculateEaster(year As Integer) As Date
    ' Algorithme de Computus pour calculer la date de Pâques
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
    ' Compte les jours fériés du mois
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
    ' Retourne le nom du jour férié pour une date donnée
    Dim easterDate As Date
    
    ' Jours fériés fixes
    If cellDate = DateSerial(year, 1, 1) Then GetHolidayName = "Jour de l'An": Exit Function
    If cellDate = DateSerial(year, 5, 1) Then GetHolidayName = "Fête du Travail": Exit Function
    If cellDate = DateSerial(year, 5, 8) Then GetHolidayName = "Fête de la Victoire 1945": Exit Function
    If cellDate = DateSerial(year, 7, 14) Then GetHolidayName = "Fête nationale": Exit Function
    If cellDate = DateSerial(year, 8, 15) Then GetHolidayName = "Assomption": Exit Function
    If cellDate = DateSerial(year, 11, 1) Then GetHolidayName = "Toussaint": Exit Function
    If cellDate = DateSerial(year, 11, 11) Then GetHolidayName = "Armistice": Exit Function
    If cellDate = DateSerial(year, 12, 25) Then GetHolidayName = "Noël": Exit Function
    
    ' Jours fériés mobiles basés sur Pâques
    easterDate = CalculateEaster(year)
    If cellDate = easterDate + 1 Then GetHolidayName = "Lundi de Pâques": Exit Function
    If cellDate = easterDate + 39 Then GetHolidayName = "Ascension": Exit Function
    If cellDate = easterDate + 50 Then GetHolidayName = "Lundi de Pentecôte": Exit Function
    
    GetHolidayName = ""
End Function

Function CountWorkingDays(month As Integer, year As Integer, holidays As Collection) As Integer
    ' Compte les jours travaillés du mois (lundi à vendredi, sauf jours fériés)
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
        
        ' Vérifier si c'est un jour travaillé (lundi à vendredi, col 0 à 4)
        If dayOfWeek >= 0 And dayOfWeek <= 4 Then
            ' Vérifier si ce n'est pas un jour férié
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
    ' Compte les cellules rouges (jours en congé)
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

