''' <summary>
''' Used to create a string array which represent time in a day.
''' </summary>
Public Class Hours
    ''' <summary>
    ''' Creates an array quarter hours
    ''' </summary>
    Public Shared ReadOnly Property Quarterly() As String()
        Get
            Return Range(TimeIncrement.Quarterly)
        End Get
    End Property
    ''' <summary>
    ''' Creates an array of hours
    ''' </summary>
    Public Shared ReadOnly Property Hourly() As String()
        Get
            Return Range()
        End Get
    End Property
    ''' <summary>
    ''' Creates an array of half-hours
    ''' </summary>
    Public Shared ReadOnly Property HalfHour() As String()
        Get
            Return Range(TimeIncrement.HalfHour)
        End Get
    End Property

    Public Shared Property TimeFormat() As String = "hh:mm tt"

    Public Shared Function Range(Optional pTimeIncrement As TimeIncrement = TimeIncrement.Hourly) As String()

        Dim hoursContainer As IEnumerable(Of Date) = Enumerable.Range(0, 24).Select(Function(index) (Date.MinValue.AddHours(index)))
        Dim timeList = New List(Of String)()

        For Each dateTime In hoursContainer

            timeList.Add(dateTime.ToString(TimeFormat))

            If pTimeIncrement = TimeIncrement.Quarterly Then
                timeList.Add(dateTime.AddMinutes(15).ToString(TimeFormat))
                timeList.Add(dateTime.AddMinutes(30).ToString(TimeFormat))
                timeList.Add(dateTime.AddMinutes(45).ToString(TimeFormat))
            ElseIf pTimeIncrement = TimeIncrement.HalfHour Then
                timeList.Add(dateTime.AddMinutes(30).ToString(TimeFormat))
            End If

        Next

        Return timeList.ToArray()

    End Function

End Class