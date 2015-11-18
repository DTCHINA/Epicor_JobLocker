Module Module1

    Dim var_today As Date
    Sub Main()
        var_today = Today.AddDays(2) ' Change the # of add days for mauall testing.
        'REM
        'Establish code for connecting to epicor
        Dim Days_Out As Integer
        Days_Out = 5
        CalcBusinessDays(Days_Out, Days_Out)

        Console.WriteLine("Start Date = " & var_today & " + " & Days_Out & " " & var_today.AddDays(Days_Out))
        Console.ReadKey()

    End Sub

    Function CalcBusinessDays(ByVal D1 As Integer, ByRef BizzSpan As Integer) As Decimal
        'In Comming is the stock number of days
        'Out Going is the calculated number of days including the weekends + 1 or 2 if the end date is a weekend day
        'Will need to calculate if there is a holiday in side the duration
        Dim C, HH As Integer
        C = 1
        HH = 0
        Dim WorkingDate As Date = var_today
        BizzSpan = D1

        Console.WriteLine("<> Start Date = " & WorkingDate & " Biz Days =" & BizzSpan)

        Do Until D1 = C
            Console.WriteLine("Counter = " & C & " Day = " & WorkingDate.ToString("ddd") & " biz  = " & BizzSpan)

            'Test for Sunday
            If Weekday(WorkingDate) = 1 Then
                BizzSpan = BizzSpan + 1 'Sunday
                C = C - 1
            End If

            'Tset for Saturday 
            If Weekday(WorkingDate) = 7 Then
                BizzSpan = BizzSpan + 1 'Saturday
                C = C - 1
            End If

            ' Test for Holiday
            CalcHolidays(WorkingDate, HH)
            C = C - HH
            BizzSpan = BizzSpan + HH

            WorkingDate = WorkingDate.AddDays(1)
            C = C + 1
        Loop

        BizzSpan = BizzSpan - 1
        WorkingDate = var_today.AddDays(BizzSpan)
        Console.WriteLine("Start date is " & var_today & " Number of days out = " & BizzSpan)
  
        '       If Weekday(WorkingDate) = 1 Then BizzSpan = BizzSpan + 1 'Sunday
        '       If Weekday(WorkingDate) = 7 Then BizzSpan = BizzSpan + 2 'Saturday
  
        Return BizzSpan

    End Function

    Function CalcHolidays(ByVal D2 As Date, ByRef Counter_Add As Integer) As Decimal
        
        If CDate(D2) = "2015/09/03" Then
            Counter_Add = 1
            Console.WriteLine("Holiday")
        Else : Counter_Add = 0
        End If

        Return Counter_Add

    End Function
End Module
