Module Module1

    Sub Main()
        'REM
        'Establish code for connecting to epicor
        Dim Days_Out As Integer

        Days_Out = 5
        Console.WriteLine("Days out Before     = " & Days_Out)
        CalcBusinessDays(Days_Out, Days_Out)
        Console.WriteLine("Days out After      = " & Days_Out)
        Console.ReadKey()

        CalcHolidays(Days_Out, Days_Out)
        Console.WriteLine("Days out Holiday    = " & Days_Out)

        'Comment the next line below to get it to pause.
        Console.ReadKey()

    End Sub

    Function CalcBusinessDays(ByVal D1 As Integer, ByRef BizzSpan As Integer) As Decimal
        Dim C As Integer = 0
        Dim WorkingDate As Date = Today
        BizzSpan = D1

        Do Until D1 = C

            WorkingDate = WorkingDate.AddDays(1)
            C = C + 1

            If Weekday(WorkingDate) = 7 Or Weekday(WorkingDate) = 1 Then BizzSpan = BizzSpan + 1
            Console.WriteLine("Date today = " & WorkingDate & " " & Weekday(WorkingDate) & " Biz Days " & BizzSpan)

        Loop


        Return BizzSpan
    End Function

    Function CalcHolidays(ByVal D1 As Integer, ByRef BizzSpan As Integer) As Decimal
        Dim C As Integer = 0
        Dim WorkingDate As Date = Today
        BizzSpan = D1

        Do Until D1 = C

            WorkingDate = WorkingDate.AddDays(1)
            C = C + 1

            If WorkingDate = "7/11/2015" Then BizzSpan = BizzSpan + 1
            If WorkingDate = "7/12/2015" Then BizzSpan = BizzSpan + 1
            Console.WriteLine("Date today = " & WorkingDate & " " & Weekday(WorkingDate) & " Biz Days " & BizzSpan)

        Loop


        Return BizzSpan
    End Function
End Module
