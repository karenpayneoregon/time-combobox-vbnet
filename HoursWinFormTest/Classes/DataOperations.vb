Imports System.Data.SqlClient

Public Class DataOperations
    Inherits BaseSqlServerConnection

    Public Shared Function Read() As List(Of TimeTable)
        Dim list = New List(Of TimeTable)

        Using cn As New SqlConnection(ConnectionString)
            Using cmd As New SqlCommand("SELECT id, FirstName, LastName, StartTime, EndTime FROM dbo.TimeTable;", cn)
                cn.Open()
                Dim reader As SqlDataReader = cmd.ExecuteReader()
                While reader.Read()
                    list.Add(New TimeTable() With
                                {
                                    .Id = reader.GetInt32(0),
                                    .FirstName = reader.GetSqlString(1),
                                    .LastName = reader.GetSqlString(2),
                                    .StartTime = reader.GetTimeSpan(3),
                                    .EndTime = reader.GetTimeSpan(4)
                                })
                End While
            End Using
        End Using

        Return list

    End Function
End Class
