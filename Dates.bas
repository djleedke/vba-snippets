Attribute VB_Name = "DateFunctions"

Function FirstDayOfMonth(selectedDate)

    '   Gets the first of the month for the provided date
    '
    '   Arguments:
    '       selectedDate: The date to find the first day of the month for.
    '
    '   Returns:
    '       date: The first day of the month for the provided date.
    
    FirstDayOfMonth = DateSerial(Year(selectedDate), Month(selectedDate), 1)

End Function


Function LastDayOfMonth(selectedDate)

    '   Returns the last day of the month for the provided date.
    '
    '   Arguments:
    '       selectedDate: The date to find the last day of the month for.
    '
    '   Returns:
    '       date: The last day of the month for the provided date.

    LastDayOfMonth = DateSerial(Year(selectedDate), Month(selectedDate) + 1, 0)

End Function


Function FirstDayOfYear(selectedDate)
    
    '   Returns the first day of the year for the provided date.
    '
    '   Arguments:
    '       selectedDate: The date to find the first day of the year for.
    '
    '   Returns:
    '       date: The first day of the year for the provided date.

    FirstDayOfYear = DateSerial(Year(selectedDate), 1, 1)

End Function


Function SameDayLastYear(selectedDate)

    '   Returns the same date a year earlier.
    '
    '   Arguments:
    '       selectedDate: The date to find the same day prior year for.
    '
    '   Returns:
    '       date: The same day last year.

    SameDayLastYear = DateSerial(Year(selectedDate) - 1, Month(selectedDate), day(selectedDate))

End Function


Function DaysBetween(startDate, endDate)

    '   Returns the number of days between two dates.
    '
    '   Arguments:
    '       startDate: The beginning date.
    '       endDate: The ending date.
    '
    '   Returns:
    '       date: The number of days between the start date and end date.

    Dim firstDate As Date, secondDate As Date, n As Integer
    DaysBetween = DateDiff("d", startDate, endDate)
    
End Function
