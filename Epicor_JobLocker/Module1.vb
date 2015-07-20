Module Module1

    Dim var_today As Date
    Sub Main()
        var_today = Today.AddDays(0) ' Change the # of add days for mauall testing.
        'REM
        'Establish code for connecting to epicor
        Dim Days_Out As Integer
        Days_Out = 5
        CalcBusinessDays(Days_Out, Days_Out)

        'Console.WriteLine("Start Date = " & var_today & " + " & Days_Out & " " & var_today.AddDays(Days_Out))
        Console.ReadKey()

    End Sub

    Function CalcBusinessDays(ByVal D1 As Integer, ByRef BizzSpan As Integer) As Decimal
        'In Comming is the stock number of days
        'Out Going is the calculated number of days including the weekends + 1 or 2 if the end date is a weekend day
        'Will need to calculate if there is a holiday in side the duration
        Dim C As Integer = 1
        Dim varTRAP As Boolean
        Dim WorkingDate As Date = var_today
        BizzSpan = D1

        Console.WriteLine("*Start Date = " & WorkingDate & " Biz Days =" & BizzSpan)
        Console.WriteLine("")

        Do Until D1 = C
            Console.WriteLine("  Date = " & WorkingDate & " days out = " & BizzSpan)
            If Weekday(WorkingDate) = 1 Then
                BizzSpan = BizzSpan + 1 'Sunday
                Console.WriteLine(" *Sunday + 1 = " & WorkingDate)
            End If

            If Weekday(WorkingDate) = 7 Then
                BizzSpan = BizzSpan + 1 'Saturday
                Console.WriteLine(" *Saturday + 1 = " & WorkingDate)
            End If

            Holiday_test(WorkingDate, varTRAP)
            If varTRAP = True Then
                BizzSpan = BizzSpan + 1
                Console.WriteLine(" *HOLIDAY + 1")
            End If

            'BizzSpan = BizzSpan + 1
            WorkingDate = WorkingDate.AddDays(1)
            C = C + 1
            varTRAP = 0
        Loop

        Console.WriteLine("*End of Loop1")
        BizzSpan = BizzSpan - 1
        WorkingDate = var_today.AddDays(BizzSpan)
        Console.WriteLine("*Start date is " & var_today & " Number of days out = " & BizzSpan)
        Console.WriteLine("that would make the end date " & WorkingDate)
        'TEST the end date

        If Weekday(WorkingDate) = 1 Then BizzSpan = BizzSpan + 1 'Sunday
        If Weekday(WorkingDate) = 7 Then BizzSpan = BizzSpan + 2 'Saturday
        Console.WriteLine("  Today " & var_today & " + " & BizzSpan & " " & var_today.AddDays(BizzSpan))

        Return BizzSpan

    End Function

    Function CalcHolidays(ByVal D1 As Integer, ByRef BizzSpan As Integer) As Decimal
        Dim C As Integer = 0
        Dim WorkingDate As Date = var_today
        BizzSpan = D1

        Do Until D1 = C
            If CDate(WorkingDate) = "2015/07/23" Then BizzSpan = BizzSpan + 1
            WorkingDate = WorkingDate.AddDays(1)
            C = C + 1
        Loop

        Return BizzSpan

    End Function
    Function Holiday_test(ByVal D1 As Date, ByRef varTrap As Boolean) As Decimal
        varTrap = False
        If D1 = "7/21/2015" Then varTrap = True
        Return varTrap
    End Function
End Module
