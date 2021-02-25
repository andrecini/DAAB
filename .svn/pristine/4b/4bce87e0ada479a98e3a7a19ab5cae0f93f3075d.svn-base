Imports System.Data.Common
''' <summary>
''' Decorator for the connection class, exposing additional info like it's transaction.
''' </summary>
Public Class ConnectionWithExtraInfo
    Inherits DbConnection
    Private m_connection As DbConnection = Nothing
    Private m_transaction As IDbTransaction = Nothing

    Public ReadOnly Property Connection() As IDbConnection
        Get
            Return m_connection
        End Get
    End Property

    Public ReadOnly Property Transaction() As IDbTransaction
        Get
            Return m_transaction
        End Get
    End Property

    Public Sub New(ByVal connection As DbConnection)
        Me.m_connection = connection
    End Sub

    Protected Overrides Function BeginDbTransaction(ByVal isolationLevel As System.Data.IsolationLevel) As System.Data.Common.DbTransaction

        If m_transaction Is Nothing Then
            m_transaction = New NestedDbTransaction(m_connection.BeginTransaction(isolationLevel), False)
        Else
            If m_transaction.Connection Is Nothing Then
                m_transaction = New NestedDbTransaction(m_connection.BeginTransaction(isolationLevel), False)
            Else
                Return New NestedDbTransaction(CType(m_transaction, NestedDbTransaction).Transaction, True)
            End If
        End If

        Return m_transaction

    End Function

#Region "Default DbConnection"

    Public Overrides Sub ChangeDatabase(ByVal databaseName As String)
        m_connection.ChangeDatabase(databaseName)
    End Sub

    Public Overrides Sub Close()
        m_connection.Close()
    End Sub


    Public Overrides Sub Open()
        m_connection.Open()
    End Sub

    Public Overrides Property ConnectionString() As String
        Get
            Return m_connection.ConnectionString
        End Get
        Set(ByVal value As String)
            m_connection.ConnectionString = value
        End Set
    End Property

    Public Overrides ReadOnly Property Database() As String
        Get
            Return m_connection.Database
        End Get
    End Property

    Public Overrides ReadOnly Property DataSource() As String
        Get
            Return m_connection.DataSource
        End Get
    End Property

    Public Overrides ReadOnly Property ServerVersion() As String
        Get
            Return m_connection.ServerVersion
        End Get
    End Property

    Public Overrides ReadOnly Property State() As System.Data.ConnectionState
        Get
            Return m_connection.State
        End Get
    End Property

    Protected Overrides Function CreateDbCommand() As System.Data.Common.DbCommand
        Return m_connection.CreateCommand()
    End Function

    Public Overrides Function GetSchema() As System.Data.DataTable
        Return m_connection.GetSchema()
    End Function

    Public Overrides Function GetSchema(ByVal collectionName As String) As System.Data.DataTable
        Return m_connection.GetSchema(collectionName)
    End Function

    Public Overrides Function GetSchema(ByVal collectionName As String, ByVal restrictionValues() As String) As System.Data.DataTable
        Return m_connection.GetSchema(collectionName, restrictionValues)
    End Function


#End Region

End Class
