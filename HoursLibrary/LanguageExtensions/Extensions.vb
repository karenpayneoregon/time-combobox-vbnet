Namespace LanguageExtensions
    Public Module Extensions
        ''' <summary>
        ''' Combine date and time
        ''' </summary>
        ''' <param name="day">Valid Initialized DateTime</param>
        ''' <param name="time">Valid initialized TimeSpan</param>
        ''' <returns>Day with Time</returns>
        <Runtime.CompilerServices.Extension>
        Public Function At(day As Date, time As TimeSpan) As Date
            Return (New DateTime(day.Year, day.Month, day.Day, 0, 0, 0)).Add(time)
        End Function

        <Runtime.CompilerServices.Extension>
        Public Function RoundedQuarterly(sender As Date) As Date
            Return New DateTime(sender.Year, sender.Month, sender.Day, sender.Hour, (sender.Minute \ 15) * 15, 0)
        End Function

        <Runtime.CompilerServices.Extension>
        Public Function RoundedUpHour(sender As Date) As Date
            Return Date.Parse($"{IIf(sender.Minute > 30, sender.AddHours(1), sender):yyyy-MM-dd HH:00:00}")
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
        ''' Format a TimeSpan hh:mm tt
        ''' </summary>
        ''' <param name="sender"></param>
        ''' <returns></returns>
        <Runtime.CompilerServices.Extension>
        Public Function CurrentHour(sender As TimeSpan) As String

            Dim time = New TimeSpan(0, sender.Hours, 0, 0)
            Dim TimeIn12Hours As String

            TimeIn12Hours = Date.MinValue.
                AddHours(time.Hours).
                AddMinutes(time.Minutes).
                ToString("hh:mm tt")

            Return TimeIn12Hours

        End Function

        ''' <summary>
        ''' Returns the original <see cref="DateTime"/> with Hour part changed to supplied hour parameter.
        ''' </summary>
        <Runtime.CompilerServices.Extension>
        Public Function SetTime(originalDate As Date, hour As Integer) As Date
            Return New DateTime(originalDate.Year, originalDate.Month, originalDate.Day, hour, originalDate.Minute, originalDate.Second, originalDate.Millisecond, originalDate.Kind)
        End Function

        ''' <summary>
        ''' Returns the original <see cref="DateTime"/> with Hour and Minute parts changed to supplied hour and minute parameters.
        ''' </summary>
        <Runtime.CompilerServices.Extension>
        Public Function SetTime(originalDate As Date, hour As Integer, minute As Integer) As Date
            Return New DateTime(originalDate.Year, originalDate.Month, originalDate.Day, hour, minute, originalDate.Second, originalDate.Millisecond, originalDate.Kind)
        End Function

        ''' <summary>
        ''' Returns the original <see cref="DateTime"/> with Hour, Minute and Second parts changed to supplied hour, minute and second parameters.
        ''' </summary>
        <Runtime.CompilerServices.Extension>
        Public Function SetTime(originalDate As Date, hour As Integer, minute As Integer, second As Integer) As Date
            Return New DateTime(originalDate.Year, originalDate.Month, originalDate.Day, hour, minute, second, originalDate.Millisecond, originalDate.Kind)
        End Function

    End Module

End Namespace