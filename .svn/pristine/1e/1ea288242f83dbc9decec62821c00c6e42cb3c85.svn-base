Option Explicit On 
Option Strict On

Imports System.Text.RegularExpressions

Public MustInherit Class DbHelper
    Implements IDBHelper

    Private m_DBHelperParameterCache As IDBHelperParameterCache
    Private m_QuotePrefix As String

    Private m_QuoteSuffix As String

    Private m_CommandTimeOut As Integer

    Private m_LowerCaseColumNames As Boolean


    'functions to create specific .net provider data objects  
    Protected MustOverride Function NewCommand() As IDbCommand
    Protected MustOverride Function NewDataAdapter(ByRef cmd As IDbCommand) As IDataAdapter

    Protected MustOverride Function NewDBHelperParameterCache() As IDBHelperParameterCache

    Protected MustOverride Function SQLParamReplaceString() As String

    'Protected MustOverride Function GetSpParameterSet(ByVal connectionString As String, ByVal spName As String, Optional ByVal includeReturnValueParameter As Boolean = False) As IDataParameter()
    Private Function GetSpParameterSet(ByVal connectionString As String, ByVal spName As String, Optional ByVal includeReturnValueParameter As Boolean = False) As IDataParameter()
        Return m_DBHelperParameterCache.GetSpParameterSet(connectionString, spName, includeReturnValueParameter)
    End Function

    Sub New()
        m_DBHelperParameterCache = NewDBHelperParameterCache()
        m_QuotePrefix = ""
        m_QuoteSuffix = ""
        m_CommandTimeOut = -1
    End Sub


    ReadOnly Property DBHelperParameterCache() As IDBHelperParameterCache Implements IDBHelper.DBHelperParameterCache
        Get
            DBHelperParameterCache = m_DBHelperParameterCache
        End Get
    End Property

    Public Property QuotePrefix() As String
        Get
            Return m_QuotePrefix
        End Get
        Set(ByVal Value As String)
            m_QuotePrefix = Value
        End Set
    End Property

    Public Property QuoteSuffix() As String
        Get
            Return m_QuoteSuffix
        End Get
        Set(ByVal Value As String)
            m_QuoteSuffix = Value
        End Set
    End Property

    Public Property CommandTimeOut() As Integer
        Get
            Return m_CommandTimeOut
        End Get
        Set(ByVal Value As Integer)
            m_CommandTimeOut = Value
        End Set
    End Property

    Public Overridable Function Null2Number(ByVal value As Object) As Object Implements IDBHelper.Null2Number
        If Convert.IsDBNull(value) Then
            Null2Number = CShort(0)
        Else
            Null2Number = value
        End If
    End Function

    Public Overridable Function Null2String(ByVal value As Object) As Object Implements IDBHelper.Null2String
        If Convert.IsDBNull(value) Then
            Null2String = ""
        Else
            Null2String = value
        End If
    End Function

    Public Property LowerCaseColumNames() As Boolean
        Get
            Return m_LowerCaseColumNames
        End Get
        Set(ByVal Value As Boolean)
            m_LowerCaseColumNames = Value
        End Set
    End Property

    ''' alternative code  http://www.codeproject.com/KB/vb/ConvToSqlDbType.aspx
    Public Overridable Function GetDbType(ByVal [type] As Type) As DbType Implements IDBHelper.GetDbType
        Dim name As String = [type].Name
        Dim val As DbType = DbType.String
        'Try
        val = CType([Enum].Parse(GetType(DbType), name, True), DbType)
        'Catch ex As Exception
        'End Try

        Return val

    End Function


#Region "Create Database Objects"
    Public MustOverride Function NewConnection(Optional ByVal sConnectionString As String = "") As IDbConnection Implements IDBHelper.NewConnection
    Public MustOverride Overloads Function NewParameter() As IDataParameter Implements IDBHelper.NewParameter
    Public MustOverride Overloads Function NewParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType) As IDataParameter Implements IDBHelper.NewParameter
    Public MustOverride Overloads Function NewParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType, ByVal size As Integer) As IDataParameter Implements IDBHelper.NewParameter
    Public MustOverride Overloads Function NewParameter(ByVal parameterName As String, ByVal value As Object) As IDataParameter Implements IDBHelper.NewParameter
    Public MustOverride Overloads Function NewParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType, ByVal value As Object) As IDataParameter Implements IDBHelper.NewParameter
#End Region

#Region "ExecuteNonQuery"

    Public Overridable Overloads Function ExecuteNonQuery(ByVal connectionString As String, _
                                                  ByVal commandType As CommandType, _
                                                  ByVal commandText As String) As Integer Implements IDBHelper.ExecuteNonQuery

        'pass through the call providing null for the set of SqlParameters
        Return ExecuteNonQuery(connectionString, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function

    Public Overridable Overloads Function ExecuteNonQuery(ByVal connectionString As String, _
                                                    ByVal commandType As CommandType, _
                                                    ByVal commandText As String, _
                                                    ByVal ParamArray commandParameters() As IDataParameter) As Integer Implements IDBHelper.ExecuteNonQuery

        Dim cn As IDbConnection = NewConnection(connectionString)

        Try
            cn.Open()

            'call the overload that takes a connection in place of the connection string
            Return ExecuteNonQuery(cn, commandType, commandText, commandParameters)
        Finally
            cn.Dispose()
        End Try

    End Function

    Public Overridable Overloads Function ExecuteNonQuery(ByVal connection As IDbConnection, _
                                                     ByVal commandType As CommandType, _
                                                     ByVal commandText As String, _
                                                     ByVal ParamArray commandParameters() As IDataParameter) As Integer Implements IDBHelper.ExecuteNonQuery


        If IsRunningTransaction(connection) Then
            Return ExecuteNonQuery(CurrentTransaction(connection), commandType, commandText, commandParameters)
        End If

        'create a command and prepare it for execution
        Dim cmd As IDbCommand = NewCommand()
        Dim retval As Integer
        Try

            PrepareCommand(cmd, connection, CType(Nothing, IDbTransaction), commandType, commandText, commandParameters)

            'finally, execute the command.
            retval = cmd.ExecuteNonQuery()

            'detach the SqlParameters from the command object, so they can be used again
            ClearParameters(cmd)

            Return retval

        Catch ex As Exception
            Throw New ApplicationException(ex.Message & " SQL:" & cmd.CommandText, ex)
        End Try

    End Function

    Public Overridable Overloads Function ExecuteNonQuery(ByVal connectionString As String, _
                                                     ByVal spName As String, _
                                                     ByVal ParamArray parameterValues() As Object) As Integer Implements IDBHelper.ExecuteNonQuery
        Dim commandParameters As IDataParameter()

        'if we receive parameter values, we need to figure out where they go
        If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
            'pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)

            commandParameters = GetSpParameterSet(connectionString, spName)

            'assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'call the overload that takes an array of SqlParameters
            Return ExecuteNonQuery(connectionString, CommandType.StoredProcedure, spName, commandParameters)
            'otherwise we can just call the SP without params
        Else
            Return ExecuteNonQuery(connectionString, CommandType.StoredProcedure, spName)
        End If

    End Function

    Public Overridable Overloads Function ExecuteNonQuery(ByVal connection As IDbConnection, _
                                                         ByVal commandType As CommandType, _
                                                         ByVal commandText As String) As Integer Implements IDBHelper.ExecuteNonQuery
        'pass through the call providing null for the set of SqlParameters
        Return ExecuteNonQuery(connection, commandType, commandText, CType(Nothing, IDataParameter()))

    End Function 'ExecuteNonQuery

    Public Overridable Overloads Function ExecuteNonQuery(ByVal connection As IDbConnection, _
                                                        ByVal spName As String, _
                                                        ByVal ParamArray parameterValues() As Object) As Integer Implements IDBHelper.ExecuteNonQuery
        Dim commandParameters As IDataParameter()

        'if we receive parameter values, we need to figure out where they go
        If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
            'pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            commandParameters = GetSpParameterSet(connection.ConnectionString, spName)

            'assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'call the overload that takes an array of SqlParameters
            Return ExecuteNonQuery(connection, CommandType.StoredProcedure, spName, commandParameters)
            'otherwise we can just call the SP without params
        Else
            Return ExecuteNonQuery(connection, CommandType.StoredProcedure, spName)
        End If

    End Function 'ExecuteNonQuery

    Public Overridable Overloads Function ExecuteNonQuery(ByVal transaction As IDbTransaction, _
                                                          ByVal commandType As CommandType, _
                                                          ByVal commandText As String) As Integer Implements IDBHelper.ExecuteNonQuery
        'pass through the call providing null for the set of SqlParameters
        Return ExecuteNonQuery(transaction, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function 'ExecuteNonQuery

    Public Overridable Overloads Function ExecuteNonQuery(ByVal transaction As IDbTransaction, _
                                                         ByVal commandType As CommandType, _
                                                         ByVal commandText As String, _
                                                         ByVal ParamArray commandParameters() As IDataParameter) As Integer Implements IDBHelper.ExecuteNonQuery
        'create a command and prepare it for execution
        Dim cmd As IDbCommand = NewCommand()
        Dim retval As Integer
        Try

            PrepareCommand(cmd, transaction.Connection, transaction, commandType, commandText, commandParameters)

            'finally, execute the command.
            retval = cmd.ExecuteNonQuery()

            'detach the SqlParameters from the command object, so they can be used again
            ClearParameters(cmd)

            Return retval

        Catch ex As Exception
#If DEBUG Then
            Dim i As Integer
            For i = 0 To commandParameters.Length - 1
                System.Diagnostics.Debug.WriteLine(commandParameters(i).ParameterName & "-" & commandParameters(i).DbType.ToString & "-" & commandParameters(i).Value.ToString)
            Next
#End If
            Throw New ApplicationException(ex.Message & " SQL:" & cmd.CommandText, ex)
        End Try

    End Function 'ExecuteNonQuery

    Public Overridable Overloads Function ExecuteNonQuery(ByVal transaction As IDbTransaction, _
                                                     ByVal spName As String, _
                                                     ByVal ParamArray parameterValues() As Object) As Integer Implements IDBHelper.ExecuteNonQuery
        Dim commandParameters As IDataParameter()

        'if we receive parameter values, we need to figure out where they go
        If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
            'pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            commandParameters = GetSpParameterSet(transaction.Connection.ConnectionString, spName)

            'assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'call the overload that takes an array of SqlParameters
            Return ExecuteNonQuery(transaction, CommandType.StoredProcedure, spName, commandParameters)
            'otherwise we can just call the SP without params
        Else
            Return ExecuteNonQuery(transaction, CommandType.StoredProcedure, spName)
        End If
    End Function 'ExecuteNonQuery

#End Region

#Region "ExecuteDataset"


    Public Overridable Overloads Function ExecuteDataset(ByVal connectionString As String, _
                                                       ByVal commandType As CommandType, _
                                                       ByVal commandText As String) As DataSet Implements IDBHelper.ExecuteDataset
        'pass through the call providing null for the set of SqlParameters
        Return ExecuteDataset(connectionString, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function 'ExecuteDataset

    Public Overridable Overloads Function ExecuteDataset(ByVal connectionString As String, _
                                                       ByVal commandType As CommandType, _
                                                       ByVal commandText As String, _
                                                       ByVal ParamArray commandParameters() As IDataParameter) As DataSet Implements IDBHelper.ExecuteDataset
        'create & open a SqlConnection, and dispose of it after we are done.
        Dim cn As IDbConnection = NewConnection(connectionString)
        Try
            cn.Open()

            'call the overload that takes a connection in place of the connection string
            Return ExecuteDataset(cn, commandType, commandText, commandParameters)
        Finally
            cn.Dispose()
        End Try
    End Function 'ExecuteDataset

    Public Overridable Overloads Function ExecuteDataset(ByVal connectionString As String, _
                                                       ByVal spName As String, _
                                                       ByVal ParamArray parameterValues() As Object) As DataSet Implements IDBHelper.ExecuteDataset

        Dim commandParameters As IDataParameter()

        'if we receive parameter values, we need to figure out where they go
        If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
            'pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            commandParameters = GetSpParameterSet(connectionString, spName)

            'assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'call the overload that takes an array of SqlParameters
            Return ExecuteDataset(connectionString, CommandType.StoredProcedure, spName, commandParameters)
            'otherwise we can just call the SP without params
        Else
            Return ExecuteDataset(connectionString, CommandType.StoredProcedure, spName)
        End If
    End Function 'ExecuteDataset

    Public Overridable Overloads Function ExecuteDataset(ByVal connection As IDbConnection, _
                                                         ByVal commandType As CommandType, _
                                                        ByVal commandText As String) As DataSet Implements IDBHelper.ExecuteDataset

        'pass through the call providing null for the set of SqlParameters
        Return ExecuteDataset(connection, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function 'ExecuteDataset

    Public Overridable Overloads Function ExecuteDataset(ByVal connection As IDbConnection, _
                                                      ByVal commandType As CommandType, _
                                                      ByVal commandText As String, _
                                                      ByVal ParamArray commandParameters() As IDataParameter) As DataSet Implements IDBHelper.ExecuteDataset


        If IsRunningTransaction(connection) Then
            Return ExecuteDataset(CurrentTransaction(connection), commandType, commandText, commandParameters)
        End If


        'create a command and prepare it for execution
        Dim cmd As IDbCommand = NewCommand()
        Dim ds As New DataSet
        Dim da As IDataAdapter

        Try

            PrepareCommand(cmd, connection, CType(Nothing, IDbTransaction), commandType, commandText, commandParameters)

            'create the DataAdapter & DataSet
            da = NewDataAdapter(cmd)

            'fill the DataSet using default values for DataTable names, etc.
            da.Fill(ds)

            'detach the SqlParameters from the command object, so they can be used again
            ClearParameters(cmd)

            If Me.LowerCaseColumNames Then
                LowerCaseDsColums(ds)
            End If


            'return the dataset
            Return ds

        Catch ex As Exception
            Throw New ApplicationException(ex.Message & " SQL:" & cmd.CommandText, ex)
        End Try

    End Function 'ExecuteDataset

    Private Sub LowerCaseDsColums(ByRef ds As DataSet)
        'lower case variables 
        Dim i As Integer
        Dim j As Integer
        For i = 0 To ds.Tables.Count - 1

            For j = 0 To ds.Tables(i).Columns.Count - 1
                ds.Tables(i).Columns(j).ColumnName = ds.Tables(i).Columns(j).ColumnName.ToLower
            Next j
        Next i

    End Sub

    Public Overridable Overloads Function ExecuteDataset(ByVal connection As IDbConnection, _
                                                        ByVal spName As String, _
                                                        ByVal ParamArray parameterValues() As Object) As DataSet Implements IDBHelper.ExecuteDataset

        'Return ExecuteDataset(connection, spName, parameterValues)
        Dim commandParameters As IDataParameter()

        'if we receive parameter values, we need to figure out where they go
        If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
            'pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            commandParameters = GetSpParameterSet(connection.ConnectionString, spName)

            'assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'call the overload that takes an array of SqlParameters
            Return ExecuteDataset(connection, CommandType.StoredProcedure, spName, commandParameters)
            'otherwise we can just call the SP without params
        Else
            Return ExecuteDataset(connection, CommandType.StoredProcedure, spName)
        End If

    End Function 'ExecuteDataset

    Public Overridable Overloads Function ExecuteDataset(ByVal transaction As IDbTransaction, _
                                                      ByVal commandType As CommandType, _
                                                      ByVal commandText As String, _
                                                      ByVal ParamArray commandParameters() As IDataParameter) As DataSet Implements IDBHelper.ExecuteDataset
        'create a command and prepare it for execution
        Dim cmd As IDbCommand = NewCommand()
        Dim ds As New DataSet
        Dim da As IDataAdapter
        Try

            PrepareCommand(cmd, transaction.Connection, transaction, commandType, commandText, commandParameters)


            'create the DataAdapter & DataSet
            da = NewDataAdapter(cmd)

            'fill the DataSet using default values for DataTable names, etc.
            da.Fill(ds)


            'detach the SqlParameters from the command object, so they can be used again
            ClearParameters(cmd)

            If Me.LowerCaseColumNames Then
                LowerCaseDsColums(ds)
            End If

            'return the dataset
            Return ds

        Catch ex As Exception
            Throw New ApplicationException(ex.Message & " SQL:" & cmd.CommandText, ex)
        End Try
    End Function 'ExecuteDataset


    Public Overridable Overloads Function ExecuteDataset(ByVal transaction As IDbTransaction, _
                                                      ByVal commandType As CommandType, _
                                                      ByVal commandText As String) As DataSet Implements IDBHelper.ExecuteDataset
        'pass through the call providing null for the set of SqlParameters
        Return ExecuteDataset(transaction, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function 'ExecuteDataset



    Public Overridable Overloads Function ExecuteDataset(ByVal transaction As IDbTransaction, _
                                                        ByVal spName As String, _
                                                        ByVal ParamArray parameterValues() As Object) As DataSet Implements IDBHelper.ExecuteDataset
        Dim commandParameters As IDataParameter()

        'if we receive parameter values, we need to figure out where they go
        If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
            'pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            commandParameters = GetSpParameterSet(transaction.Connection.ConnectionString, spName)

            'assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'call the overload that takes an array of SqlParameters
            Return ExecuteDataset(transaction, CommandType.StoredProcedure, spName, commandParameters)
            'otherwise we can just call the SP without params
        Else
            Return ExecuteDataset(transaction, CommandType.StoredProcedure, spName)
        End If
    End Function 'ExecuteDataset





#End Region

#Region "ExecuteReader"
    ' this enum is used to indicate whether the connection was provided by the caller, or created by SqlHelper, so that
    ' we can set the appropriate CommandBehavior when calling ExecuteReader()
    Private Enum SqlConnectionOwnership
        'Connection is owned and managed by SqlHelper
        Internal
        'Connection is owned and managed by the caller
        [External]
    End Enum 'SqlConnectionOwnership

    ' Create and prepare a SqlCommand, and call ExecuteReader with the appropriate CommandBehavior.
    ' If we created and opened the connection, we want the connection to be closed when the DataReader is closed.
    ' If the caller provided the connection, we want to leave it to them to manage.
    ' Parameters:
    ' -connection - a valid SqlConnection, on which to execute this command 
    ' -transaction - a valid SqlTransaction, or 'null' 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' -commandParameters - an array of SqlParameters to be associated with the command or 'null' if no parameters are required 
    ' -connectionOwnership - indicates whether the connection parameter was provided by the caller, or created by SqlHelper 
    ' Returns: SqlDataReader containing the results of the command 
    Private Overloads Function ExecuteReader(ByVal connection As IDbConnection, _
                                                    ByVal transaction As IDbTransaction, _
                                                    ByVal commandType As CommandType, _
                                                    ByVal commandText As String, _
                                                    ByVal commandParameters() As IDataParameter, _
                                                    ByVal connectionOwnership As SqlConnectionOwnership) As IDataReader
        'create a command and prepare it for execution
        Dim cmd As IDbCommand = NewCommand()
        'create a reader
        Dim dr As IDataReader

        PrepareCommand(cmd, connection, transaction, commandType, commandText, commandParameters)

        ' call ExecuteReader with the appropriate CommandBehavior
        If connectionOwnership = SqlConnectionOwnership.External Then
            dr = cmd.ExecuteReader()
        Else
            dr = cmd.ExecuteReader(CommandBehavior.CloseConnection)
        End If

        ClearParameters(cmd)

        Return dr
    End Function 'ExecuteReader

    Public Overridable Overloads Function ExecuteReader(ByVal connectionString As String, _
                                                      ByVal commandType As CommandType, _
                                                      ByVal commandText As String) As IDataReader Implements IDBHelper.ExecuteReader
        'pass through the call providing null for the set of SqlParameters
        Return ExecuteReader(connectionString, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function 'ExecuteReader

    Public Overridable Overloads Function ExecuteReader(ByVal connectionString As String, _
                                                      ByVal commandType As CommandType, _
                                                      ByVal commandText As String, _
                                                      ByVal ParamArray commandParameters() As IDataParameter) As IDataReader Implements IDBHelper.ExecuteReader
        'create & open a SqlConnection
        Dim cn As IDbConnection = NewConnection(connectionString)
        cn.Open()

        Try
            'call the private overload that takes an internally owned connection in place of the connection string
            Return ExecuteReader(cn, CType(Nothing, IDbTransaction), commandType, commandText, commandParameters, SqlConnectionOwnership.Internal)
        Catch ex As Exception
            'if we fail to return the SqlDatReader, we need to close the connection ourselves
            System.Diagnostics.Debug.WriteLine(ex.Message)
            cn.Dispose()
        End Try
    End Function 'ExecuteReader

    Public Overridable Overloads Function ExecuteReader(ByVal connectionString As String, _
                                                       ByVal spName As String, _
                                                       ByVal ParamArray parameterValues() As Object) As IDataReader Implements IDBHelper.ExecuteReader
        Dim commandParameters As IDataParameter()

        'if we receive parameter values, we need to figure out where they go
        If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
            'pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            commandParameters = GetSpParameterSet(connectionString, spName)

            'assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'call the overload that takes an array of SqlParameters
            Return ExecuteReader(connectionString, CommandType.StoredProcedure, spName, commandParameters)
            'otherwise we can just call the SP without params
        Else
            Return ExecuteReader(connectionString, CommandType.StoredProcedure, spName)
        End If
    End Function 'ExecuteReader

    Public Overridable Overloads Function ExecuteReader(ByVal connection As IDbConnection, _
                                                          ByVal commandType As CommandType, _
                                                          ByVal commandText As String) As IDataReader Implements IDBHelper.ExecuteReader

        Return ExecuteReader(connection, commandType, commandText, CType(Nothing, IDataParameter()))

    End Function 'ExecuteReader

    Public Overridable Overloads Function ExecuteReader(ByVal connection As IDbConnection, _
                                                      ByVal commandType As CommandType, _
                                                      ByVal commandText As String, _
                                                      ByVal ParamArray commandParameters() As IDataParameter) As IDataReader Implements IDBHelper.ExecuteReader


        If IsRunningTransaction(connection) Then
            Return ExecuteReader(CurrentTransaction(connection), commandType, commandText, commandParameters)
        End If

        'pass through the call to private overload using a null transaction value
        Return ExecuteReader(connection, CType(Nothing, IDbTransaction), commandType, commandText, commandParameters, SqlConnectionOwnership.External)

    End Function 'ExecuteReader

    Public Overridable Overloads Function ExecuteReader(ByVal connection As IDbConnection, _
                                                      ByVal spName As String, _
                                                      ByVal ParamArray parameterValues() As Object) As IDataReader Implements IDBHelper.ExecuteReader
        'pass through the call using a null transaction value
        'Return ExecuteReader(connection, CType(Nothing, SqlTransaction), spName, parameterValues)

        Dim commandParameters As IDataParameter()

        'if we receive parameter values, we need to figure out where they go
        If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
            commandParameters = GetSpParameterSet(connection.ConnectionString, spName)

            AssignParameterValues(commandParameters, parameterValues)

            Return ExecuteReader(connection, CommandType.StoredProcedure, spName, commandParameters)

            'otherwise we can just call the SP without params
        Else
            Return ExecuteReader(connection, CommandType.StoredProcedure, spName)
        End If

    End Function 'ExecuteReader

    Public Overridable Overloads Function ExecuteReader(ByVal transaction As IDbTransaction, _
                                                      ByVal commandType As CommandType, _
                                                      ByVal commandText As String) As IDataReader Implements IDBHelper.ExecuteReader
        'pass through the call providing null for the set of SqlParameters
        Return ExecuteReader(transaction, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function 'ExecuteReader

    Public Overridable Overloads Function ExecuteReader(ByVal transaction As IDbTransaction, _
                                                       ByVal commandType As CommandType, _
                                                       ByVal commandText As String, _
                                                       ByVal ParamArray commandParameters() As IDataParameter) As IDataReader Implements IDBHelper.ExecuteReader
        'pass through to private overload, indicating that the connection is owned by the caller
        Return ExecuteReader(transaction.Connection, transaction, commandType, commandText, commandParameters, SqlConnectionOwnership.External)
    End Function 'ExecuteReader

    Public Overridable Overloads Function ExecuteReader(ByVal transaction As IDbTransaction, _
                                                          ByVal spName As String, _
                                                          ByVal ParamArray parameterValues() As Object) As IDataReader Implements IDBHelper.ExecuteReader
        Dim commandParameters As IDataParameter()

        'if we receive parameter values, we need to figure out where they go
        If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
            commandParameters = GetSpParameterSet(transaction.Connection.ConnectionString, spName)

            AssignParameterValues(commandParameters, parameterValues)

            Return ExecuteReader(transaction, CommandType.StoredProcedure, spName, commandParameters)
            'otherwise we can just call the SP without params
        Else
            Return ExecuteReader(transaction, CommandType.StoredProcedure, spName)
        End If
    End Function 'ExecuteReader






#End Region

#Region "ExecuteScalar"
    Public Overridable Overloads Function ExecuteScalar(ByVal connectionString As String, _
                                                          ByVal commandType As CommandType, _
                                                          ByVal commandText As String) As Object Implements IDBHelper.ExecuteScalar
        'pass through the call providing null for the set of SqlParameters
        Return ExecuteScalar(connectionString, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function 'ExecuteScalar

    Public Overridable Overloads Function ExecuteScalar(ByVal connectionString As String, _
                                                       ByVal commandType As CommandType, _
                                                       ByVal commandText As String, _
                                                       ByVal ParamArray commandParameters() As IDataParameter) As Object Implements IDBHelper.ExecuteScalar
        'create & open a SqlConnection, and dispose of it after we are done.
        Dim cn As IDbConnection = NewConnection(connectionString)
        Try
            cn.Open()

            'call the overload that takes a connection in place of the connection string
            Return ExecuteScalar(cn, commandType, commandText, commandParameters)
        Finally
            cn.Dispose()
        End Try
    End Function 'ExecuteScalar

    Public Overridable Overloads Function ExecuteScalar(ByVal connectionString As String, _
                                                      ByVal spName As String, _
                                                      ByVal ParamArray parameterValues() As Object) As Object Implements IDBHelper.ExecuteScalar
        Dim commandParameters As IDataParameter()

        'if we receive parameter values, we need to figure out where they go
        If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
            'pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            commandParameters = GetSpParameterSet(connectionString, spName)

            'assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'call the overload that takes an array of SqlParameters
            Return ExecuteScalar(connectionString, CommandType.StoredProcedure, spName, commandParameters)
            'otherwise we can just call the SP without params
        Else
            Return ExecuteScalar(connectionString, CommandType.StoredProcedure, spName)
        End If
    End Function 'ExecuteScalar

    Public Overridable Overloads Function ExecuteScalar(ByVal connection As IDbConnection, _
                                                       ByVal commandType As CommandType, _
                                                       ByVal commandText As String) As Object Implements IDBHelper.ExecuteScalar
        Dim commandParameters As IDataParameter()
        'pass through the call providing null for the set of SqlParameters
        Return ExecuteScalar(connection, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function 'ExecuteScalar

    Public Overridable Overloads Function ExecuteScalar(ByVal connection As IDbConnection, _
                                                       ByVal commandType As CommandType, _
                                                       ByVal commandText As String, _
                                                       ByVal ParamArray commandParameters() As IDataParameter) As Object Implements IDBHelper.ExecuteScalar

        If IsRunningTransaction(connection) Then
            Return ExecuteScalar(CurrentTransaction(connection), commandType, commandText, commandParameters)
        End If

        'create a command and prepare it for execution
        Dim cmd As IDbCommand = NewCommand()
        Dim retval As Object

        Try


            PrepareCommand(cmd, connection, CType(Nothing, IDbTransaction), commandType, commandText, commandParameters)

            'execute the command & return the results
            retval = cmd.ExecuteScalar()

            'detach the SqlParameters from the command object, so they can be used again
            ClearParameters(cmd)

            Return retval

        Catch ex As Exception
            Throw New ApplicationException(ex.Message & " SQL:" & cmd.CommandText, ex)
        End Try

    End Function 'ExecuteScalar


    Public Overridable Overloads Function ExecuteScalar(ByVal connection As IDbConnection, _
                                                     ByVal spName As String, _
                                                     ByVal ParamArray parameterValues() As Object) As Object Implements IDBHelper.ExecuteScalar

        Dim commandParameters() As IDataParameter

        'if we receive parameter values, we need to figure out where they go
        If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
            'pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            commandParameters = GetSpParameterSet(connection.ConnectionString, spName)

            'assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'call the overload that takes an array of SqlParameters
            Return ExecuteScalar(connection, CommandType.StoredProcedure, spName, commandParameters)
            'otherwise we can just call the SP without params
        Else
            Return ExecuteScalar(connection, CommandType.StoredProcedure, spName)
        End If

    End Function 'ExecuteScalar

    Public Overridable Overloads Function ExecuteScalar(ByVal transaction As IDbTransaction, _
                                                      ByVal commandType As CommandType, _
                                                      ByVal commandText As String) As Object Implements IDBHelper.ExecuteScalar
        'pass through the call providing null for the set of SqlParameters
        Return ExecuteScalar(transaction, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function 'ExecuteScalar

    Public Overridable Overloads Function ExecuteScalar(ByVal transaction As IDbTransaction, _
                                                      ByVal commandType As CommandType, _
                                                      ByVal commandText As String, _
                                                      ByVal ParamArray commandParameters() As IDataParameter) As Object Implements IDBHelper.ExecuteScalar
        'create a command and prepare it for execution
        Dim cmd As IDbCommand = NewCommand()
        Dim retval As Object
        Try


            PrepareCommand(cmd, transaction.Connection, transaction, commandType, commandText, commandParameters)

            'execute the command & return the results
            retval = cmd.ExecuteScalar()

            'detach the SqlParameters from the command object, so they can be used again
            ClearParameters(cmd)

            Return retval

        Catch ex As Exception
            Throw New ApplicationException(ex.Message & " SQL:" & cmd.CommandText, ex)
        End Try

    End Function 'ExecuteScalar

    Public Overridable Overloads Function ExecuteScalar(ByVal transaction As IDbTransaction, _
                                                      ByVal spName As String, _
                                                      ByVal ParamArray parameterValues() As Object) As Object Implements IDBHelper.ExecuteScalar
        Dim commandParameters() As IDataParameter

        'if we receive parameter values, we need to figure out where they go
        If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
            'pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            commandParameters = GetSpParameterSet(transaction.Connection.ConnectionString, spName)

            'assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'call the overload that takes an array of SqlParameters
            Return ExecuteScalar(transaction, CommandType.StoredProcedure, spName, commandParameters)
            'otherwise we can just call the SP without params
        Else
            Return ExecuteScalar(transaction, CommandType.StoredProcedure, spName)
        End If
    End Function 'ExecuteScalar

#End Region

#Region "UpdateDataset"

    Public Overridable Overloads Function UpdateDataset(ByVal connection As IDbConnection, _
                                                    ByVal commandText As String, ByVal ds As DataSet, ByVal srcTable As String) As Integer Implements IDBHelper.UpdateDataset

        Dim table As DataTable

        table = ds.Tables(srcTable)

        Return UpdateDataset(connection, commandText, table)

    End Function

    Public Overridable Overloads Function UpdateDataset(ByVal transcation As IDbTransaction, _
                                                   ByVal commandText As String, ByVal ds As DataSet, ByVal srcTable As String) As Integer Implements IDBHelper.UpdateDataset
        Dim table As DataTable

        table = ds.Tables(srcTable)

        Return UpdateDataset(transcation, commandText, table)

    End Function

    Public Overridable Overloads Function UpdateDataset(ByVal connection As IDbConnection, _
                                       ByVal ds As DataSet, _
                                       ByVal srcTable As String, _
                                       Optional ByVal UpdateCommandText As String = "", _
                                       Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                       Optional ByVal InsertCommandText As String = "", _
                                       Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                       Optional ByVal DeleteCommandText As String = "", _
                                       Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer Implements IDBHelper.UpdateDataset


        Dim table As DataTable

        table = ds.Tables(srcTable)

        Return UpdateDataset(connection, table, UpdateCommandText, UpdatedataParam, InsertCommandText, InsertdataParam, DeleteCommandText, DeletedataParam)

    End Function

    Public Overridable Overloads Function UpdateDataset(ByVal transaction As IDbTransaction, _
                                           ByVal ds As DataSet, _
                                           ByVal srcTable As String, _
                                           Optional ByVal UpdateCommandText As String = "", _
                                           Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                           Optional ByVal InsertCommandText As String = "", _
                                           Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                           Optional ByVal DeleteCommandText As String = "", _
                                           Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer Implements IDBHelper.UpdateDataset
        Dim table As DataTable

        table = ds.Tables(srcTable)

        Return UpdateDataset(transaction, table, UpdateCommandText, UpdatedataParam, InsertCommandText, InsertdataParam, DeleteCommandText, DeletedataParam)


    End Function

    Public Overridable Overloads Function UpdateDataset(ByVal connection As IDbConnection, _
                                                        ByVal commandText As String, ByVal Table As DataTable) As Integer Implements IDBHelper.UpdateDataset

        Throw New Exception("UpdateDataset not implemented")

    End Function

    Public Overridable Overloads Function UpdateDataset(ByVal transaction As IDbTransaction, _
                                                       ByVal commandText As String, ByVal Table As DataTable) As Integer Implements IDBHelper.UpdateDataset

        Throw New Exception("UpdateDataset not implemented")

    End Function

    Public Overridable Overloads Function UpdateDataset(ByVal connection As IDbConnection, _
                                       ByVal Table As DataTable, _
                                       Optional ByVal UpdateCommandText As String = "", _
                                       Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                       Optional ByVal InsertCommandText As String = "", _
                                       Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                       Optional ByVal DeleteCommandText As String = "", _
                                       Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer Implements IDBHelper.UpdateDataset

        Throw New Exception("UpdateDataset not implemented with these parameters")

    End Function

    Public Overridable Overloads Function UpdateDataset(ByVal connection As IDbTransaction, _
                                       ByVal Table As DataTable, _
                                       Optional ByVal UpdateCommandText As String = "", _
                                       Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                       Optional ByVal InsertCommandText As String = "", _
                                       Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                       Optional ByVal DeleteCommandText As String = "", _
                                       Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer Implements IDBHelper.UpdateDataset

        Throw New Exception("UpdateDataset not implemented with these parameters")

    End Function

    Public Overridable Overloads Function UpdateDataset(ByVal connection As IDbConnection, _
                                      ByVal Table As DataTable, _
                                      ByVal UpdateCommandType As CommandType, _
                                      ByVal InsertCommandType As CommandType, _
                                      ByVal DeleteCommandType As CommandType, _
                                      Optional ByVal UpdateCommand As String = "", _
                                      Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                      Optional ByVal InsertCommand As String = "", _
                                      Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                      Optional ByVal DeleteCommand As String = "", _
                                      Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer Implements IDBHelper.UpdateDataset

        Throw New Exception("UpdateDataset not implemented with these parameters")

    End Function

    Public Overridable Overloads Function UpdateDataset(ByVal transaction As IDbTransaction, _
                                      ByVal Table As DataTable, _
                                      ByVal UpdateCommandType As CommandType, _
                                      ByVal InsertCommandType As CommandType, _
                                      ByVal DeleteCommandType As CommandType, _
                                      Optional ByVal UpdateCommand As String = "", _
                                      Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                      Optional ByVal InsertCommand As String = "", _
                                      Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                      Optional ByVal DeleteCommand As String = "", _
                                      Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer Implements IDBHelper.UpdateDataset

        Throw New Exception("UpdateDataset not implemented with these parameters")
    End Function

    Public Overridable Overloads Function UpdateDataset(ByVal connection As IDbConnection, _
                                 ByVal ds As DataSet, _
                                 ByVal srcTable As String, _
                                 ByVal UpdateCommandType As CommandType, _
                                 ByVal InsertCommandType As CommandType, _
                                 ByVal DeleteCommandType As CommandType, _
                                 Optional ByVal UpdateCommand As String = "", _
                                 Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                 Optional ByVal InsertCommand As String = "", _
                                 Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                 Optional ByVal DeleteCommand As String = "", _
                                 Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer Implements IDBHelper.UpdateDataset


        Dim table As DataTable

        table = ds.Tables(srcTable)

        Return UpdateDataset(connection, table, UpdateCommandType, InsertCommandType, DeleteCommandType, UpdateCommand, UpdatedataParam, InsertCommand, InsertdataParam, DeleteCommand, DeletedataParam)

    End Function

    Public Overridable Overloads Function UpdateDataset(ByVal transaction As IDbTransaction, _
                                 ByVal ds As DataSet, _
                                 ByVal srcTable As String, _
                                 ByVal UpdateCommandType As CommandType, _
                                 ByVal InsertCommandType As CommandType, _
                                 ByVal DeleteCommandType As CommandType, _
                                 Optional ByVal UpdateCommand As String = "", _
                                 Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                 Optional ByVal InsertCommand As String = "", _
                                 Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                 Optional ByVal DeleteCommand As String = "", _
                                 Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer Implements IDBHelper.UpdateDataset


        Dim table As DataTable

        table = ds.Tables(srcTable)

        Return UpdateDataset(transaction, table, UpdateCommandType, InsertCommandType, DeleteCommandType, UpdateCommand, UpdatedataParam, InsertCommand, InsertdataParam, DeleteCommand, DeletedataParam)

    End Function

#End Region


    Public Function FormatSQLParameters(ByVal commandText As String, ByVal ParamArray commandParameters() As IDataParameter) As String



        If (commandParameters Is Nothing) Then
            Return commandText
        End If

        Dim objRegExParameters As Regex = New Regex("@\w+")
        Dim objMatches As MatchCollection
        Dim parameterMismatch As Boolean = False
        Dim paramExceptionMsg As String

        objMatches = objRegExParameters.Matches(commandText)

        Dim p As IDataParameter
        Dim i As Integer


        If commandParameters.Length <> objMatches.Count Then
            parameterMismatch = True

        Else

            For i = 0 To commandParameters.Length - 1
                If objMatches(i).Value.Trim <> commandParameters(i).ParameterName.Trim Then
                    parameterMismatch = True
                    Exit For
                End If
            Next i
        End If

        If parameterMismatch Then

            Dim CompareTable(0, 1) As String
            Dim MaxSize As Integer
            Dim FoundPoint As Boolean

            MaxSize = Math.Max(objMatches.Count - 1, commandParameters.Length - 1)
            ReDim CompareTable(MaxSize, 1)

            'Get SQL Parameters and passed parameters 
            'paramExceptionMsg = " SQL Parameters:"
            For i = 0 To objMatches.Count - 1
                'throw Error Message 
                CompareTable(i, 0) = objMatches(i).Value
                'paramExceptionMsg &= objMatches(i).Value & ";" & vbCrLf
            Next i

            'paramExceptionMsg &= " command Parameters:"

            For i = 0 To commandParameters.Length - 1
                'throw Error Message 
                CompareTable(i, 1) = commandParameters(i).ParameterName
            Next i

            paramExceptionMsg &= "SQL Parameters  Command Parameters" & vbCrLf

            FoundPoint = False

            For i = 0 To MaxSize - 1
                paramExceptionMsg &= CompareTable(i, 0) & ";" & CompareTable(i, 1)

                If (CompareTable(i, 0).Trim <> CompareTable(i, 1).Trim) And Not FoundPoint Then
                    FoundPoint = True
                    paramExceptionMsg &= " <-- ***Error***" & vbCrLf
                Else
                    paramExceptionMsg &= vbCrLf
                End If
            Next

            paramExceptionMsg &= vbCrLf & vbCrLf & commandText

            System.Diagnostics.Debug.WriteLine(commandText)
            System.Diagnostics.Debug.WriteLine(paramExceptionMsg)

            Throw New ArgumentException(paramExceptionMsg)

        End If

        'Replace @paramN with ? for OleDB and ODBC Providers  
        Dim sReplaceParm As String = SQLParamReplaceString()

        If sReplaceParm <> "" Then
            If sReplaceParm = ":" Then
                commandText = commandText.Replace("@", sReplaceParm)
                commandText = commandText.Replace("[", """")
                commandText = commandText.Replace("]", """")
            Else
                commandText = objRegExParameters.Replace(commandText, sReplaceParm)
            End If

        End If

        Return commandText

    End Function

#Region "protected  utility methods & constructors"

    ' This method is used to attach array of SqlParameters to a SqlCommand.
    ' This method will assign a value of DbNull to any parameter with a direction of
    ' InputOutput and a value of null.  
    ' This behavior will prevent default values from being used, but
    ' this will be the less common case than an intended pure output parameter (derived as InputOutput)
    ' where the user provided no input value.
    ' Parameters:
    ' -command - The command to which the parameters will be added
    ' -commandParameters - an array of SqlParameters tho be added to command
    Protected Sub AttachParameters(ByVal command As IDbCommand, ByVal commandParameters() As IDataParameter)
        Dim p As IDataParameter
        For Each p In commandParameters
            'check for derived output value with no value assigned
            If p.Direction = ParameterDirection.InputOutput And p.Value Is Nothing Then
                p.Value = Nothing
            End If
            command.Parameters.Add(p)
        Next p
    End Sub 'AttachParameters

    ' This method assigns an array of values to an array of SqlParameters.
    ' Parameters:
    ' -commandParameters - array of SqlParameters to be assigned values
    ' -array of objects holding the values to be assigned
    Protected Sub AssignParameterValues(ByVal commandParameters() As IDataParameter, ByVal parameterValues() As Object)

        Dim i As Integer
        Dim j As Integer

        If (commandParameters Is Nothing) And (parameterValues Is Nothing) Then
            'do nothing if we get no data
            Return
        End If

        ' we must have the same number of values as we pave parameters to put them in
        If commandParameters.Length <> parameterValues.Length Then
            Throw New ArgumentException("Parameter count does not match Parameter Value count.")
        End If

        'value array
        j = commandParameters.Length - 1
        For i = 0 To j
            commandParameters(i).Value = parameterValues(i)
        Next

    End Sub 'AssignParameterValues

    ' This method opens (if necessary) and assigns a connection, transaction, command type and parameters 
    ' to the provided command.
    ' Parameters:
    ' -command - the SqlCommand to be prepared
    ' -connection - a valid SqlConnection, on which to execute this command
    ' -transaction - a valid SqlTransaction, or 'null'
    ' -commandType - the CommandType (stored procedure, text, etc.)
    ' -commandText - the stored procedure name or T-SQL command
    ' -commandParameters - an array of SqlParameters to be associated with the command or 'null' if no parameters are required
    Protected Overridable Sub PrepareCommand(ByVal command As IDbCommand, _
                                      ByVal connection As IDbConnection, _
                                      ByVal transaction As IDbTransaction, _
                                      ByVal commandType As CommandType, _
                                      ByVal commandText As String, _
                                      ByVal commandParameters() As IDataParameter)


        'verify SQL Paramters from Sting and change @pn to ? for odbc and oledb 
        If commandType = commandType.Text Then
            commandText = FormatSQLParameters(commandText, commandParameters)
        End If

        'if the provided connection is not open, we will open it
        If connection.State <> ConnectionState.Open Then
            connection.Open()
        End If

        'associate the connection with the command
        command.Connection = GetAdoDotNetDbConnection(connection)


        'set the command text (stored procedure name or SQL statement)
        command.CommandText = commandText

        'if we were provided a transaction, assign it.
        If Not (transaction Is Nothing) Then
            command.Transaction = GetAdoDotNetDbTransaction(transaction)
        End If

        'set the command type
        command.CommandType = commandType

        'attach the command parameters if they are provided
        If Not (commandParameters Is Nothing) Then
            AttachParameters(command, commandParameters)
        End If

        Return
    End Sub 'PrepareCommand

    Protected Overridable Sub ClearParameters(ByRef cmd As IDbCommand)
        If Not cmd Is Nothing Then
            Try
                cmd.Parameters.Clear()
                cmd.Dispose()
            Catch ex As Exception
                System.Diagnostics.Debug.WriteLine("DBHelper:ClearParameters:" & ex.Message)
            End Try
        End If
    End Sub

#End Region


#Region "extra Info Connection"

    Protected Function GetAdoDotNetDbTransaction(ByVal transaction As IDbTransaction) As IDbTransaction

        If TypeOf transaction Is NestedDbTransaction Then
            Return CType(transaction, NestedDbTransaction).Transaction
        Else
            Return transaction
        End If
    End Function



    Protected Function GetAdoDotNetDbConnection(ByVal connection As IDbConnection) As IDbConnection

        If TypeOf connection Is ConnectionWithExtraInfo Then
            Return CType(connection, ConnectionWithExtraInfo).Connection
        Else
            Return connection
        End If
    End Function

    Protected Function IsConnectionWithExtraInfo(ByVal connection As IDbConnection) As Boolean

        If TypeOf connection Is ConnectionWithExtraInfo Then
            Return True
        Else
            Return False
        End If

    End Function


    Protected Function IsRunningTransaction(ByVal connection As IDbConnection) As Boolean
        If IsConnectionWithExtraInfo(connection) Then
            Dim connExtra = CType(connection, ConnectionWithExtraInfo)
            If Not connExtra.Transaction Is Nothing Then
                Dim trans = connExtra.Transaction
                If trans.Connection Is Nothing Then
                    Return False
                Else
                    Return True
                End If

            End If
        End If
        Return False
    End Function

    Protected Function CurrentTransaction(ByVal connection As IDbConnection) As IDbTransaction
        If IsConnectionWithExtraInfo(connection) Then
            Return CType(connection, ConnectionWithExtraInfo).Transaction
        End If
        Return Nothing
    End Function

    Protected Function GetConnectionTransaction(ByVal connection As IDbConnection) As IDbTransaction

        Dim dbTransaction As IDbTransaction = Nothing
        If IsRunningTransaction(connection) Then
            dbTransaction = CurrentTransaction(connection)
        End If
        Return dbTransaction
    End Function

#End Region


End Class
