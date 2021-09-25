Public MustInherit Class BaseSqlServerConnection
    Inherits BaseExceptionsHandler
    ''' <summary>
    ''' This points to your database server
    ''' </summary>
    Protected Shared DatabaseServer As String = ".\SQLEXPRESS"
    ''' <summary>
    ''' Name of database containing required tables
    ''' </summary>
    Protected Shared DefaultCatalog As String = "DateTimeDatabase"
    Public Shared ReadOnly Property ConnectionString() As String
        Get
            Return $"Data Source={DatabaseServer};Initial Catalog={DefaultCatalog};Integrated Security=True"
        End Get
    End Property
End Class