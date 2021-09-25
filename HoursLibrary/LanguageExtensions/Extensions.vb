Namespace LanguageExtensions
    Public Module Extensions

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

        <Runtime.CompilerServices.Extension>
        Public Function CurrentHour(sender As TimeSpan) As String
            Dim time = New TimeSpan(0, sender.Hours, 0, 0)

            Dim TimeIn12Hours As String
            TimeIn12Hours = Date.MinValue.AddHours(time.Hours).AddMinutes(time.Minutes).ToString("hh:mm tt")
            Return TimeIn12Hours
        End Function


    End Module

End Namespace