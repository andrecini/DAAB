Option Explicit On 
Option Strict On

Imports System.Data.Odbc

Public Class ODBCDBHelper
    Inherits DbHelper
    
    Public Sub New()

    End Sub

    Public Overrides Function NewConnection(Optional ByVal sConnectionString As String = "") As IDbConnection

        If sConnectionString = "" Then
            NewConnection = New OdbcConnection()
        Else
            NewConnection = New OdbcConnection(sConnectionString)
        End If

    End Function

    Public Overloads Overrides Function NewParameter() As IDataParameter
        NewParameter = New OdbcParameter()
    End Function

    Public Overloads Overrides Function NewParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType) As IDataParameter
        NewParameter = New OdbcParameter()
        NewParameter.DbType = dbType
        NewParameter.ParameterName = parameterName
    End Function
    Public Overloads Overrides Function NewParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType, ByVal size As Integer) As IDataParameter
        Dim objNewParameter As New OdbcParameter()
        objNewParameter.DbType = dbType
        objNewParameter.ParameterName = parameterName
        objNewParameter.Size = size
        NewParameter = objNewParameter
    End Function

    Public Overloads Overrides Function NewParameter(ByVal parameterName As String, ByVal value As Object) As IDataParameter
        NewParameter = New OdbcParameter(parameterName, value)
    End Function

    Public Overloads Overrides Function NewParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType, ByVal value As Object) As IDataParameter
        Dim objNewParameter As New OdbcParameter()
        objNewParameter.ParameterName = parameterName
        objNewParameter.DbType = dbType
        objNewParameter.Value = value
        NewParameter = CType(objNewParameter, IDataParameter)
    End Function


    Protected Overrides Function NewCommand() As IDbCommand
        NewCommand = New OdbcCommand()
    End Function

    Protected Overrides Function NewDataAdapter(ByRef cmd As IDbCommand) As IDataAdapter
        NewDataAdapter = New OdbcDataAdapter(CType(cmd, OdbcCommand))
    End Function

    Protected Overrides Function NewDBHelperParameterCache() As IDBHelperParameterCache
        NewDBHelperParameterCache = New ODBCDBHelperParameterCache()
    End Function

    ' This method opens (if necessary) and assigns a connection, transaction, command type and parameters 
    ' to the provided command.
    ' Parameters:
    ' -command - the SqlCommand to be prepared
    ' -connection - a valid SqlConnection, on which to execute this command
    ' -transaction - a valid SqlTransaction, or 'null'
    ' -commandType - the CommandType (stored procedure, text, etc.)
    ' -commandText - the stored procedure name or T-SQL command
    ' -commandParameters - an array of SqlParameters to be associated with the command or 'null' if no parameters are required
    Protected Overrides Sub PrepareCommand(ByVal command As IDbCommand, _
                                      ByVal connection As IDbConnection, _
                                      ByVal transaction As IDbTransaction, _
                                      ByVal commandType As CommandType, _
                                      ByVal commandText As String, _
                                      ByVal commandParameters() As IDataParameter)

        'verify SQL Paramters from Sting and change @pn to ? for odbc and OleDB 
        If commandType = commandType.Text Then
            commandText = FormatSQLParameters(commandText, commandParameters)
        End If

        'if the provided connection is not open, we will open it
        If connection.State <> ConnectionState.Open Then
            connection.Open()
        End If

        'associate the connection with the command
        command.Connection = connection

        'set the command text (stored procedure name or SQL statement)
        command.CommandText = commandText

        'if we were provided a transaction, assign it.
        If Not (transaction Is Nothing) Then
            command.Transaction = transaction
        End If

        'set the command type
        command.CommandType = commandType

        'attach the command parameters if they are provided
        If Not (commandParameters Is Nothing) Then
            AttachParameters(command, commandParameters)
        End If

        If command.CommandType = commandType.StoredProcedure Then
            If Not (commandParameters Is Nothing) Then
                command.CommandText = CreateOdbcSpSyntax(commandText, commandParameters.Length)
            Else
                command.CommandText = CreateOdbcSpSyntax(commandText, 0)
            End If
        End If

        Return
    End Sub 'PrepareCommand


    Private Function CreateOdbcSpSyntax(ByVal sProcedureName As String, ByVal iNumParam As Integer) As String

        Dim i As Integer
        Dim sODBCName As String
        sProcedureName = sProcedureName.Trim()

        If sProcedureName.StartsWith("{") Then 'probaly obdc syntax do nothing 
            CreateOdbcSpSyntax = sProcedureName
            Exit Function
        End If

        sODBCName = "{ call " & sProcedureName & "("
        For i = 1 To iNumParam
            If i > 1 Then
                sODBCName &= ","
            End If
            sODBCName &= "?"
        Next

        sODBCName &= ") }"

        Return sODBCName

    End Function

#Region "UpdateDataset"

    Private Overloads Function UpdateDatasetTransaction(ByVal connection As IDbConnection, ByVal transaction As IDbTransaction, _
                                                      ByVal commandText As String, ByVal Table As DataTable) As Integer
        'trwo  error messages 

        Dim oDA As IDataAdapter
        Dim cdm As IDbCommand

        cdm = NewCommand()
        PrepareCommand(cdm, connection, transaction, CommandType.Text, commandText, Nothing)
        oDA = NewDataAdapter(cdm)

        Dim oCB As OdbcCommandBuilder = New OdbcCommandBuilder(CType(oDA, OdbcDataAdapter))

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


        Dim oDA As OdbcDataAdapter
        oDA = New OdbcDataAdapter()
        Dim cdm As IDbCommand

        'Update Command
        If UpdateCommandText.Trim() <> "" Then
            cdm = NewCommand()
            PrepareCommand(cdm, connection, transaction, UpdateCommandType, UpdateCommandText, UpdatedataParam)
            oDA.UpdateCommand = CType(cdm, OdbcCommand)
        End If

        'Insert  Command 
        If InsertCommandText.Trim() <> "" Then
            cdm = NewCommand()
            PrepareCommand(cdm, connection, transaction, InsertCommandType, InsertCommandText, InsertdataParam)
            oDA.InsertCommand = CType(cdm, OdbcCommand)
        End If

        'Delete  Command 
        If DeleteCommandText.Trim() <> "" Then
            cdm = NewCommand()
            PrepareCommand(cdm, connection, transaction, DeleteCommandType, DeleteCommandText, DeletedataParam)
            oDA.DeleteCommand = CType(cdm, OdbcCommand)
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
                                                      ByVal commandText As String, ByVal Table As DataTable) As Integer

        Return UpdateDatasetTransaction(connection, Nothing, commandText, Table)

    End Function

    Public Overloads Overrides Function UpdateDataset(ByVal transaction As IDbTransaction, _
                                                    ByVal commandText As String, ByVal Table As DataTable) As Integer

        Return UpdateDatasetTransaction(transaction.Connection, transaction, commandText, Table)

    End Function


    Public Overloads Overrides Function UpdateDataset(ByVal connection As IDbConnection, _
                                       ByVal Table As DataTable, _
                                       Optional ByVal UpdateCommandText As String = "", _
                                       Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                       Optional ByVal InsertCommandText As String = "", _
                                       Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                       Optional ByVal DeleteCommandText As String = "", _
                                       Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer



        Return UpdateDatasetTransaction(connection, Nothing, Table, _
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

        Return UpdateDatasetTransaction(connection, Nothing, Table, _
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



    Protected Overrides Function SQLParamReplaceString() As String
        SQLParamReplaceString = "?"
    End Function

#End Region

End Class


' SqlHelperParameterCache provides functions to leverage a static cache of procedure parameters, and the
' ability to discover parameters for stored procedures at run-time.
Public NotInheritable Class ODBCDBHelperParameterCache
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

        Dim cn As New OdbcConnection(connectionString)
        Dim cmd As OdbcCommand = New OdbcCommand(spName, cn)
        Dim discoveredParameters() As IDataParameter

        Try
            cn.Open()
            cmd.CommandType = CommandType.StoredProcedure
            OdbcCommandBuilder.DeriveParameters(cmd)
            If Not includeReturnValueParameter Then
                cmd.Parameters.RemoveAt(0)
            End If

            discoveredParameters = New OdbcParameter(cmd.Parameters.Count - 1) {}
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
        Dim cachedParameters() As IDataParameter = CType(paramCache(hashKey), IDataParameter())

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

        hashKey = connectionString + ":" + spName + CStr(IIf(includeReturnValueParameter = True, ":include ReturnValue Parameter", ""))

        cachedParameters = CType(paramCache(hashKey), IDataParameter())

        If (cachedParameters Is Nothing) Then
            paramCache(hashKey) = DiscoverSpParameterSet(connectionString, spName, includeReturnValueParameter)
            cachedParameters = CType(paramCache(hashKey), IDataParameter())

        End If

        Return CloneParameters(cachedParameters)

    End Function 'GetSpParameterSet
#End Region

End Class 'SqlHelperParameterCache 

