Option Explicit On 
Option Strict On

Imports System.Data.SqlClient

Public Class SqlDbHelper
    Inherits DbHelper

    Public Sub New()

    End Sub

    Public Overrides Function NewConnection(Optional ByVal sConnectionString As String = "") As IDbConnection

        If sConnectionString = "" Then
            NewConnection = New ConnectionWithExtraInfo(New SqlConnection())
        Else
            NewConnection = New ConnectionWithExtraInfo(New SqlConnection(sConnectionString))
        End If

    End Function

    Public Overloads Overrides Function NewParameter() As IDataParameter
        NewParameter = New SqlClient.SqlParameter()
    End Function

    Public Overloads Overrides Function NewParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType) As IDataParameter
        NewParameter = NewParameter()
        NewParameter.ParameterName = parameterName
        NewParameter.DbType = dbType
    End Function
    Public Overloads Overrides Function NewParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType, ByVal size As Integer) As IDataParameter
        Dim objSqlParameter As New SqlParameter()
        objSqlParameter.Size = size
        objSqlParameter.DbType = dbType
        objSqlParameter.ParameterName = parameterName
        NewParameter = CType(objSqlParameter, IDataParameter)
    End Function

    Public Overloads Overrides Function NewParameter(ByVal parameterName As String, ByVal value As Object) As IDataParameter
        NewParameter = New SqlClient.SqlParameter(parameterName, value)
    End Function

    Public Overloads Overrides Function NewParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType, ByVal value As Object) As IDataParameter
        Dim objNewParameter As New SqlParameter()
        objNewParameter.ParameterName = parameterName

        If dbType = dbType.[SByte] Then
            objNewParameter.DbType = dbType.Byte
        Else
            objNewParameter.DbType = dbType
        End If

        objNewParameter.Value = value
        NewParameter = CType(objNewParameter, IDataParameter)
    End Function

    Protected Overrides Function NewCommand() As IDbCommand
        NewCommand = New SqlClient.SqlCommand
        If Me.CommandTimeOut >= 0 Then
            NewCommand.CommandTimeout = Me.CommandTimeOut
        End If
    End Function

    Protected Overrides Function NewDataAdapter(ByRef cmd As IDbCommand) As IDataAdapter
        NewDataAdapter = New SqlDataAdapter(CType(cmd, SqlCommand))
    End Function


    Protected Overrides Function NewDBHelperParameterCache() As IDBHelperParameterCache
        NewDBHelperParameterCache = New SqlHelperParameterCache()
    End Function

#Region "ExecuteXmlReader"

    ' Execute a SqlCommand (that returns a resultset and takes no parameters) against the provided SqlConnection. 
    ' e.g.:  
    ' Dim r As XmlReader = ExecuteXmlReader(conn, CommandType.StoredProcedure, "GetOrders")
    ' Parameters:
    ' -connection - a valid SqlConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command using "FOR XML AUTO" 
    ' Returns: an XmlReader containing the resultset generated by the command 
    Public Overloads Function ExecuteXmlReader(ByVal connection As IDbConnection, _
                                                      ByVal commandType As CommandType, _
                                                      ByVal commandText As String) As XmlReader
        'pass through the call providing null for the set of SqlParameters
        Return ExecuteXmlReader(connection, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function 'ExecuteXmlReader

    ' Execute a SqlCommand (that returns a resultset) against the specified SqlConnection 
    ' using the provided parameters.
    ' e.g.:  
    ' Dim r As XmlReader = ExecuteXmlReader(conn, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))
    ' Parameters:
    ' -connection - a valid SqlConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command using "FOR XML AUTO" 
    ' -commandParameters - an array of SqlParamters used to execute the command 
    ' Returns: an XmlReader containing the resultset generated by the command 
    Public Overloads Function ExecuteXmlReader(ByVal connection As IDbConnection, _
                                                      ByVal commandType As CommandType, _
                                                      ByVal commandText As String, _
                                                      ByVal ParamArray commandParameters() As IDataParameter) As XmlReader
        'pass through the call using a null transaction value
        'Return ExecuteXmlReader(connection, CType(Nothing, SqlTransaction), commandType, commandText, commandParameters)
        'create a command and prepare it for execution
        Dim cmd As New SqlCommand()
        Dim retval As XmlReader

        PrepareCommand(cmd, connection, CType(Nothing, IDbTransaction), commandType, commandText, commandParameters)

        'create the DataAdapter & DataSet
        retval = cmd.ExecuteXmlReader()

        'detach the SqlParameters from the command object, so they can be used again
        cmd.Parameters.Clear()

        Return retval


    End Function 'ExecuteXmlReader

    ' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified SqlConnection 
    ' using the provided parameter values.  This method will discover the parameters for the 
    ' stored procedure, and assign the values based on parameter order.
    ' This method provides no access to output parameters or the stored procedure's return value parameter.
    ' e.g.:  
    ' Dim r As XmlReader = ExecuteXmlReader(conn, "GetOrders", 24, 36)
    ' Parameters:
    ' -connection - a valid SqlConnection 
    ' -spName - the name of the stored procedure using "FOR XML AUTO" 
    ' -parameterValues - an array of objects to be assigned as the input values of the stored procedure 
    ' Returns: an XmlReader containing the resultset generated by the command 
    Public Overloads Function ExecuteXmlReader(ByVal connection As IDbConnection, _
                                                      ByVal spName As String, _
                                                      ByVal ParamArray parameterValues() As Object) As XmlReader
        Dim commandParameters As IDataParameter()

        'if we receive parameter values, we need to figure out where they go
        If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
            'pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            commandParameters = DBHelperParameterCache.GetSpParameterSet(connection.ConnectionString, spName)

            'assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'call the overload that takes an array of SqlParameters
            Return ExecuteXmlReader(connection, CommandType.StoredProcedure, spName, commandParameters)
            'otherwise we can just call the SP without params
        Else
            Return ExecuteXmlReader(connection, CommandType.StoredProcedure, spName)
        End If
    End Function 'ExecuteXmlReader


    ' Execute a SqlCommand (that returns a resultset and takes no parameters) against the provided SqlTransaction
    ' e.g.:  
    ' Dim r As XmlReader = ExecuteXmlReader(trans, CommandType.StoredProcedure, "GetOrders")
    ' Parameters:
    ' -transaction - a valid SqlTransaction
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command using "FOR XML AUTO" 
    ' Returns: an XmlReader containing the resultset generated by the command 
    Public Overloads Function ExecuteXmlReader(ByVal transaction As IDbTransaction, _
                                                      ByVal commandType As CommandType, _
                                                      ByVal commandText As String) As XmlReader
        'pass through the call providing null for the set of SqlParameters
        Return ExecuteXmlReader(transaction, commandType, commandText, CType(Nothing, IDataParameter()))
    End Function 'ExecuteXmlReader

    ' Execute a SqlCommand (that returns a resultset) against the specified SqlTransaction
    ' using the provided parameters.
    ' e.g.:  
    ' Dim r As XmlReader = ExecuteXmlReader(trans, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))
    ' Parameters:
    ' -transaction - a valid SqlTransaction
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command using "FOR XML AUTO" 
    ' -commandParameters - an array of SqlParamters used to execute the command 
    ' Returns: an XmlReader containing the resultset generated by the command
    Public Overloads Function ExecuteXmlReader(ByVal transaction As IDbTransaction, _
                                                      ByVal commandType As CommandType, _
                                                      ByVal commandText As String, _
                                                      ByVal ParamArray commandParameters() As IDataParameter) As XmlReader
        'create a command and prepare it for execution
        Dim cmd As IDbCommand = NewCommand()
        Dim retval As XmlReader

        PrepareCommand(cmd, transaction.Connection, transaction, commandType, commandText, commandParameters)

        'create the DataAdapter & DataSet
        retval = CType(cmd, SqlClient.SqlCommand).ExecuteXmlReader()

        'detach the SqlParameters from the command object, so they can be used again
        cmd.Parameters.Clear()

        Return retval

    End Function 'ExecuteXmlReader

    ' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified SqlTransaction 
    ' using the provided parameter values.  This method will discover the parameters for the 
    ' stored procedure, and assign the values based on parameter order.
    ' This method provides no access to output parameters or the stored procedure's return value parameter.
    ' e.g.:  
    ' Dim r As XmlReader = ExecuteXmlReader(trans, "GetOrders", 24, 36)
    ' Parameters:
    ' -transaction - a valid SqlTransaction
    ' -spName - the name of the stored procedure 
    ' -parameterValues - an array of objects to be assigned as the input values of the stored procedure 
    ' Returns: a dataset containing the resultset generated by the command
    Public Overloads Function ExecuteXmlReader(ByVal transaction As IDbTransaction, _
                                                      ByVal spName As String, _
                                                      ByVal ParamArray parameterValues() As Object) As XmlReader
        Dim commandParameters As IDataParameter()

        'if we receive parameter values, we need to figure out where they go
        If Not (parameterValues Is Nothing) And parameterValues.Length > 0 Then
            'pull the parameters for this stored procedure from the parameter cache (or discover them & populate the cache)
            commandParameters = DBHelperParameterCache.GetSpParameterSet(transaction.Connection.ConnectionString, spName)

            'assign the provided values to these parameters based on parameter order
            AssignParameterValues(commandParameters, parameterValues)

            'call the overload that takes an array of SqlParameters
            Return ExecuteXmlReader(transaction, CommandType.StoredProcedure, spName, commandParameters)
            'otherwise we can just call the SP without params
        Else
            Return ExecuteXmlReader(transaction, CommandType.StoredProcedure, spName)
        End If
    End Function 'ExecuteXmlReader

#End Region

#Region "UpdateDataset"

    Public Overloads Overrides Function UpdateDataset(ByVal connection As IDbConnection, _
                                                      ByVal commandText As String, ByVal Table As DataTable) As Integer

        Dim dbTransaction As IDbTransaction = GetConnectionTransaction(connection)

        Return UpdateDatasetTransaction(connection, dbTransaction, commandText, Table)

    End Function


    Public Overloads Overrides Function UpdateDataset(ByVal transaction As IDbTransaction, _
                                                    ByVal commandText As String, ByVal Table As DataTable) As Integer

        Return UpdateDatasetTransaction(transaction.Connection, transaction, commandText, Table)

    End Function

    Private Overloads Function UpdateDatasetTransaction(ByVal connection As IDbConnection, ByVal trans As IDbTransaction, _
                                                      ByVal commandText As String, ByVal Table As DataTable) As Integer
        'trwo  error messages 

        Dim oDA As IDataAdapter
        Dim cdm As IDbCommand


        cdm = NewCommand()
        PrepareCommand(cdm, connection, trans, CommandType.Text, commandText, Nothing)
        oDA = NewDataAdapter(cdm)

        Dim oCB As SqlCommandBuilder = New SqlCommandBuilder(CType(oDA, SqlDataAdapter))

        oCB.QuotePrefix = Me.QuotePrefix
        oCB.QuoteSuffix = Me.QuoteSuffix

        Dim rowsAffected As Integer
        Try
            rowsAffected = CType(oDA, System.Data.Common.DbDataAdapter).Update(Table)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message())
            Throw ex
        End Try

        Return rowsAffected

    End Function

    Private Overloads Function UpdateDatasetTransaction(ByVal connection As IDbConnection, _
                                     ByVal transaction As IDbTransaction, _
                                     ByVal Table As DataTable, _
                                     ByVal UpdateCommandType As CommandType, _
                                     ByVal InsertCommandType As CommandType, _
                                     ByVal DeleteCommandType As CommandType, _
                                     Optional ByVal UpdateCommandText As String = "", _
                                     Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                     Optional ByVal InsertCommandText As String = "", _
                                     Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                     Optional ByVal DeleteCommandText As String = "", _
                                     Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer


        Dim oDA As SqlDataAdapter
        oDA = New SqlDataAdapter()
        Dim cdm As IDbCommand

        'Update Command
        If UpdateCommandText.Trim() <> "" Then
            cdm = NewCommand()
            PrepareCommand(cdm, connection, transaction, UpdateCommandType, UpdateCommandText, UpdatedataParam)
            oDA.UpdateCommand = CType(cdm, SqlCommand)
        End If

        'Insert Command 
        If InsertCommandText.Trim() <> "" Then
            cdm = NewCommand()
            PrepareCommand(cdm, connection, transaction, InsertCommandType, InsertCommandText, InsertdataParam)
            oDA.InsertCommand = CType(cdm, SqlCommand)
        End If

        'Delete  Command 
        If DeleteCommandText.Trim() <> "" Then
            cdm = NewCommand()
            PrepareCommand(cdm, connection, transaction, DeleteCommandType, DeleteCommandText, DeletedataParam)
            oDA.DeleteCommand = CType(cdm, SqlCommand)
        End If


        Dim iNumRowsAffected As Integer

        With oDA
            iNumRowsAffected = .Update(Table)
            ClearParameters(CType(.SelectCommand, IDbCommand))
            ClearParameters(CType(.DeleteCommand, IDbCommand))
            ClearParameters(CType(.InsertCommand, IDbCommand))
            ClearParameters(CType(.UpdateCommand, IDbCommand))
        End With

        ClearParameters(cdm)

        Return iNumRowsAffected

    End Function


    Public Overloads Overrides Function UpdateDataset(ByVal connection As IDbConnection, _
                                       ByVal Table As DataTable, _
                                       Optional ByVal UpdateCommandText As String = "", _
                                       Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                       Optional ByVal InsertCommandText As String = "", _
                                       Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                       Optional ByVal DeleteCommandText As String = "", _
                                       Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer

         Dim dbTransaction As IDbTransaction = GetConnectionTransaction(connection)

        Return UpdateDatasetTransaction(connection, dbTransaction, Table, _
                                  CommandType.Text, CommandType.Text, CommandType.Text, _
                                  UpdateCommandText, _
                                  UpdatedataParam, _
                                  InsertCommandText, _
                                  InsertdataParam, _
                                  DeleteCommandText, _
                                  DeletedataParam)
    End Function

    Public Overloads Overrides Function UpdateDataset(ByVal transaction As IDbTransaction, _
                                   ByVal Table As DataTable, _
                                   Optional ByVal UpdateCommandText As String = "", _
                                   Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                   Optional ByVal InsertCommandText As String = "", _
                                   Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                   Optional ByVal DeleteCommandText As String = "", _
                                   Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer



        Return UpdateDatasetTransaction(transaction.Connection, transaction, Table, _
                                  CommandType.Text, CommandType.Text, CommandType.Text, _
                                  UpdateCommandText, _
                                  UpdatedataParam, _
                                  InsertCommandText, _
                                  InsertdataParam, _
                                  DeleteCommandText, _
                                  DeletedataParam)

    End Function


    Public Overloads Overrides Function UpdateDataset(ByVal connection As IDbConnection, _
                                    ByVal Table As DataTable, _
                                    ByVal UpdateCommandType As CommandType, _
                                    ByVal InsertCommandType As CommandType, _
                                    ByVal DeleteCommandType As CommandType, _
                                    Optional ByVal UpdateCommand As String = "", _
                                    Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                    Optional ByVal InsertCommand As String = "", _
                                    Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                    Optional ByVal DeleteCommand As String = "", _
                                    Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer


        Dim dbTransaction As IDbTransaction = GetConnectionTransaction(connection)

        Return UpdateDatasetTransaction(connection, dbTransaction, Table, _
                                UpdateCommandType, InsertCommandType, DeleteCommandType, _
                                UpdateCommand, _
                                UpdatedataParam, _
                                InsertCommand, _
                                InsertdataParam, _
                                DeleteCommand, _
                                DeletedataParam)

    End Function

    Public Overloads Overrides Function UpdateDataset(ByVal transaction As IDbTransaction, _
                                  ByVal Table As DataTable, _
                                  ByVal UpdateCommandType As CommandType, _
                                  ByVal InsertCommandType As CommandType, _
                                  ByVal DeleteCommandType As CommandType, _
                                  Optional ByVal UpdateCommand As String = "", _
                                  Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                  Optional ByVal InsertCommand As String = "", _
                                  Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                  Optional ByVal DeleteCommand As String = "", _
                                   Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer

        Return UpdateDatasetTransaction(transaction.Connection, transaction, Table, _
                            UpdateCommandType, InsertCommandType, DeleteCommandType, _
                            UpdateCommand, _
                            UpdatedataParam, _
                            InsertCommand, _
                            InsertdataParam, _
                            DeleteCommand, _
                            DeletedataParam)
    End Function


#End Region

    Protected Overrides Function SQLParamReplaceString() As String
        SQLParamReplaceString = "" 'no replacement for MS SQL 
    End Function

End Class


' SqlHelperParameterCache provides functions to leverage a static cache of procedure parameters, and the
' ability to discover parameters for stored procedures at run-time.
Public NotInheritable Class SqlHelperParameterCache
    Implements IDBHelperParameterCache

#Region "private methods, variables, and constructors"


    'Since this class provides only static methods, make the default constructor private to prevent 
    'instances from being created with "new SqlHelperParameterCache()".
    Friend Sub New()
    End Sub 'New 

    Private paramCache As Hashtable = Hashtable.Synchronized(New Hashtable())

    ' resolve at run time the appropriate set of SqlParameters for a stored procedure
    ' Parameters:
    ' - connectionString - a valid connection string for a SqlConnection
    ' - spName - the name of the stored procedure
    ' - includeReturnValueParameter - whether or not to include their return value parameter>
    ' Returns: SqlParameter()
    Private Function DiscoverSpParameterSet(ByVal connectionString As String, _
                                                   ByVal spName As String, _
                                                   ByVal includeReturnValueParameter As Boolean, _
                                                   ByVal ParamArray parameterValues() As Object) As IDataParameter()

        Dim cn As New SqlConnection(connectionString)
        Dim cmd As SqlCommand = New SqlCommand(spName, cn)
        Dim discoveredParameters() As IDataParameter

        Try
            cn.Open()
            cmd.CommandType = CommandType.StoredProcedure
            SqlCommandBuilder.DeriveParameters(cmd)
            If Not includeReturnValueParameter Then
                cmd.Parameters.RemoveAt(0)
            End If

            discoveredParameters = New SqlParameter(cmd.Parameters.Count - 1) {}
            cmd.Parameters.CopyTo(discoveredParameters, 0)
        Finally
            cmd.Dispose()
            cn.Dispose()

        End Try

        Return discoveredParameters

    End Function 'DiscoverSpParameterSet

    'deep copy of cached SqlParameter array
    Private Function CloneParameters(ByVal originalParameters() As IDataParameter) As IDataParameter()

        Dim i As Integer
        Dim j As Integer = originalParameters.Length - 1
        Dim clonedParameters(j) As IDataParameter

        For i = 0 To j
            clonedParameters(i) = CType(CType(originalParameters(i), ICloneable).Clone, IDataParameter)
        Next

        Return clonedParameters
    End Function 'CloneParameters



#End Region

#Region "caching functions"

    ' add parameter array to the cache
    ' Parameters
    ' -connectionString - a valid connection string for a SqlConnection 
    ' -commandText - the stored procedure name or T-SQL command 
    ' -commandParameters - an array of SqlParamters to be cached 
    Public Sub CacheParameterSet(ByVal connectionString As String, _
                                        ByVal commandText As String, _
                                        ByVal ParamArray commandParameters() As IDataParameter) Implements IDBHelperParameterCache.CacheParameterSet
        Dim hashKey As String = connectionString + ":" + commandText

        paramCache(hashKey) = commandParameters
    End Sub 'CacheParameterSet

    ' retrieve a parameter array from the cache
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection 
    ' -commandText - the stored procedure name or T-SQL command 
    ' Returns: an array of SqlParamters 
    Public Function GetCachedParameterSet(ByVal connectionString As String, ByVal commandText As String) As IDataParameter() Implements IDBHelperParameterCache.GetCachedParameterSet
        Dim hashKey As String = connectionString + ":" + commandText
        Dim cachedParameters As IDataParameter() = CType(paramCache(hashKey), IDataParameter())

        If cachedParameters Is Nothing Then
            Return Nothing
        Else
            Return CloneParameters(cachedParameters)
        End If
    End Function 'GetCachedParameterSet

#End Region

#Region "Parameter Discovery Functions"
    ' Retrieves the set of SqlParameters appropriate for the stored procedure
    ' 
    ' This method will query the database for this information, and then store it in a cache for future requests.
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection 
    ' -spName - the name of the stored procedure 
    ' Returns: an array of SqlParameters
    Public Overloads Function GetSpParameterSet(ByVal connectionString As String, ByVal spName As String) As IDataParameter() Implements IDBHelperParameterCache.GetSpParameterSet
        Return GetSpParameterSet(connectionString, spName, False)
    End Function 'GetSpParameterSet 

    ' Retrieves the set of SqlParameters appropriate for the stored procedure
    ' 
    ' This method will query the database for this information, and then store it in a cache for future requests.
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection
    ' -spName - the name of the stored procedure 
    ' -includeReturnValueParameter - a bool value indicating whether the return value parameter should be included in the results 
    ' Returns: an array of SqlParameters 
    Public Overloads Function GetSpParameterSet(ByVal connectionString As String, _
                                                       ByVal spName As String, _
                                                       ByVal includeReturnValueParameter As Boolean) As IDataParameter() Implements IDBHelperParameterCache.GetSpParameterSet

        Dim cachedParameters() As IDataParameter
        Dim hashKey As String

        hashKey = connectionString & ":" & spName & CStr(IIf(includeReturnValueParameter = True, ":include ReturnValue Parameter", ""))

        cachedParameters = CType(paramCache(hashKey), IDataParameter())

        If (cachedParameters Is Nothing) Then
            paramCache(hashKey) = DiscoverSpParameterSet(connectionString, spName, includeReturnValueParameter)
            cachedParameters = CType(paramCache(hashKey), IDataParameter())

        End If

        Return CloneParameters(cachedParameters)

    End Function 'GetSpParameterSet
#End Region

End Class 'SqlHelperParameterCache 