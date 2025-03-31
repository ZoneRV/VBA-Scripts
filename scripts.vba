Function NextTaktTime(prevTime As Date, taktTime As Double, holidays As Range) As Date
    
    Dim dayStart As Date
    dayStart = TimeValue("6:30am")
            
    ' Add takt time
    NextTaktTime = DateAdd("n", taktTime * 60, prevTime)
    
    Dim i As Integer
    
    ' Add time between breaks if nessersarry
    For i = 0 To 2
        
        startMins = TotalTimeValueMinutes(prevTime)
        timeMins = TotalTimeValueMinutes(NextTaktTime)
        stopMins = TotalTimeValueMinutes(workPeriod(i, 0))
        
        isAfterCurrentIndex = timeMins > stopMins
        isStartedAfterCurrentIndex = startMins > stopMins
    
        If (isAfterCurrentIndex And Not isStartedAfterCurrentIndex) Then
        
            NextTaktTime = DateAdd("n", TotalTimeValueMinutes(workPeriod(i, 1)), NextTaktTime)
                       
        End If
        
    Next i

    ' If the next takt time is less than the starting time or greater than the finishing time add the time between end and start times
    If (TotalTimeValueMinutes(NextTaktTime) < TotalTimeValueMinutes(dayStart) Or TotalTimeValueMinutes(NextTaktTime) > TotalTimeValueMinutes(workPeriod(2, 0))) Then
        NextTaktTime = DateAdd("n", TotalTimeValueMinutes(workPeriod(2, 1)), NextTaktTime)
    
    End If
    
    ' Keep adding days until the next takt time is not a weekend or holiday
    Do While (IsHolidayOrWeekend(NextTaktTime, holidays))
        NextTaktTime = DateAdd("d", 1, NextTaktTime)
    Loop
    
    ' Add time between breaks if nessersarry
    For i = 0 To 2
        
        startMins = TotalTimeValueMinutes(prevTime)
        timeMins = TotalTimeValueMinutes(NextTaktTime)
        stopMins = TotalTimeValueMinutes(workPeriod(i, 0))
        
        isAfterCurrentIndex = timeMins > stopMins
        isStartedAfterCurrentIndex = startMins > stopMins
    
        If (isAfterCurrentIndex And Not isStartedAfterCurrentIndex) Then
        
            NextTaktTime = DateAdd("n", TotalTimeValueMinutes(workPeriod(i, 1)), NextTaktTime)
                       
        End If
        
    Next i
    
    ' Once again check for holidays or weekends in case of day rollover
    While (IsHolidayOrWeekend(NextTaktTime, holidays))
        NextTaktTime = DateAdd("d", 1, NextTaktTime)
    Wend

    ' Dont need to check for breaks again
    
End Function


Function workPeriod(i As Integer, j As Integer) As Date

    Dim workPeriods(0 To 2, 0 To 1) As Date
        
    ' Start time for first break
    workPeriods(0, 0) = TimeValue("10:00 AM")
    
    ' duration for first break
    workPeriods(0, 1) = TimeValue("00:30 AM") ' 30m
    
    ' Start time for second break
    workPeriods(1, 0) = TimeValue("12:30 PM")
    
    ' duration for second break
    workPeriods(1, 1) = TimeValue("00:15 AM") ' 15m
    
    ' Start time for end of day
    workPeriods(2, 0) = TimeValue("2:36 PM")
    
    ' duration for end of day
    workPeriods(2, 1) = TimeValue("3:54 PM") ' 15h 54m
    
    workPeriod = workPeriods(i, j)
End Function


Function TotalTimeValueMinutes(time As Date) As Double
    TotalTimeValueMinutes = Minute(time)
    TotalTimeValueMinutes = TotalTimeValueMinutes + Hour(time) * 60
End Function


Function IsHolidayOrWeekend(day As Date, holidays As Range) As Boolean

    If (Weekday(day) = 1) Then
        IsHolidayOrWeekend = True
        
    ElseIf (Weekday(day) = 7) Then
        IsHolidayOrWeekend = True
        
    End If
    
    For Each cell In holidays
    
        If (Not IsEmpty(cell)) Then
        
            holiday = CDate(cell)
            
            If ((Int(holiday)) = (Int(day))) Then
                IsHolidayOrWeekend = True
                Exit Function
            End If
        End If
        
    Next cell
    
End Function
