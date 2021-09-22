Imports System.Globalization

Public Module Extensions
    <Runtime.CompilerServices.Extension>
    Public Function RoundedQuarterly(dateValue As Date) As Date
        Return New DateTime(dateValue.Year, dateValue.Month, dateValue.Day, dateValue.Hour, (dateValue.Minute \ 15) * 15, 0)
    End Function
    <Runtime.CompilerServices.Extension>
    Public Function RoundedUpHour(dateValue As Date) As Date
        Return Date.Parse($"{IIf(dateValue.Minute > 30, dateValue.AddHours(1), dateValue):yyyy-MM-dd HH:00:00}")
    End Function

    <Runtime.CompilerServices.Extension>
    Public Function CurrentHour(dateValue As Date) As String
        Return dateValue.ToString("HH:00 tt", CultureInfo.InvariantCulture)
    End Function

    ''' <summary>
    ''' Convert TimeSpan into DateTime
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <returns></returns>
    ''' <remarks>
    ''' Intended to be used when the date part does not matter
    ''' </remarks>
    <Runtime.CompilerServices.Extension>
    Public Function ToDateTime(sender As TimeSpan) As DateTime
        Return DateTime.ParseExact(sender.Formatted("hh:mm"), "H:mm", Nothing, DateTimeStyles.None)
    End Function

    ''' <summary>
    ''' Format a TimeSpan with AM PM
    ''' </summary>
    ''' <param name="sender">TimeSpan to format</param>
    ''' <param name="format">Optional format</param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension>
    Public Function Formatted(sender As TimeSpan, Optional format As String = "hh:mm tt") As String
        Return Date.Today.Add(sender).ToString(format)
    End Function
    ''' <summary>
    ''' Truncate to current hour
    ''' </summary>
    ''' <param name="dateValue"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension>
    Public Function TruncateToHourStart(dateValue As Date) As DateTime
        Return New DateTime(dateValue.Year, dateValue.Month, dateValue.Day, dateValue.Hour, 0, 0)
    End Function
    ''' <summary>
    ''' Truncate to current minute
    ''' </summary>
    ''' <param name="dateValue"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension>
    Public Function TruncateToMinuteStart(dateValue As Date) As DateTime
        Return New DateTime(dateValue.Year, dateValue.Month, dateValue.Day, dateValue.Hour, dateValue.Minute, 0)
    End Function

    ''' <summary>
    ''' Combine date and time
    ''' </summary>
    ''' <param name="day">Valid Initialized DateTime</param>
    ''' <param name="time">Valid initialized TimeSpan</param>
    ''' <returns>Day with Time</returns>
    <Runtime.CompilerServices.Extension>
    Public Function At(day As Date, time As TimeSpan) As DateTime
        Return (New DateTime(day.Year, day.Month, day.Day, 0, 0, 0)).Add(time)
    End Function
    ''' <summary>
    ''' Remove milliseconds from date time
    ''' </summary>
    ''' <param name="dateValue"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension>
    Public Function RemoveMilliseconds(dateValue As Date) As DateTime
        Return DateTime.ParseExact(dateValue.ToString("yyyy-MM-dd HH:mm:ss"), "yyyy-MM-dd HH:mm:ss", Nothing)
    End Function

    ''' <summary>
    ''' Remove milliseconds and seconds from date time
    ''' </summary>
    ''' <param name="dateValue"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension>
    Public Function RemoveSeconds(dateValue As DateTime) As DateTime
        Return DateTime.ParseExact(dateValue.ToString("yyyy-MM-dd HH:mm"), "yyyy-MM-dd HH:mm", Nothing)
    End Function

    ''' <summary>
    ''' Remove milliseconds and seconds from date time
    ''' </summary>
    ''' <param name="dateValue"></param>
    ''' <returns></returns>
    <Runtime.CompilerServices.Extension>
    Public Function RemoveMillisecondsAndSeconds(dateValue As Date) As DateTime
        Return DateTime.ParseExact(dateValue.ToString("yyyy-MM-dd HH:mm"), "yyyy-MM-dd HH:mm", Nothing)
    End Function
End Module

