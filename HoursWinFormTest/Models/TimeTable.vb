Imports System.Data.SqlTypes

Public Class TimeTable
    Public Property Id() As Integer
    Public Property FirstName() As SqlString
    Public Property LastName() As SqlString
    Public Property StartTime() As TimeSpan?
    Public Property EndTime() As TimeSpan?

    Public Overrides Function ToString() As String
        Return $"{FirstName} {LastName}"
    End Function
End Class
