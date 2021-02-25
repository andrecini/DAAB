Option Explicit On 
Option Strict On

Imports System.Data.OracleClient
Imports System.Text.RegularExpressions

Public Class MSOracleDbHelper
    Inherits DbHelper

    ' Requires Windows XP SP1 or Windows Server 2003.
    Private Declare Auto Function SetDllDirectory Lib "kernel32.dll" ( _
        ByVal lpPathName As String _
    ) As Boolean

    Public Function SetOracleDllDirectory(ByVal oraclePath As String) As Boolean
        Return SetDllDirectory(oraclePath)
    End Function

    Public Sub New()
        Me.LowerCaseColumNames = True
    End Sub

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

        Dim parm As IDataParameter
        Dim paramName As String
        Dim paramShort As String
        Dim i As Integer

        commandText = commandText.ToUpper

        If Not commandParameters Is Nothing Then
            For Each parm In commandParameters
                parm.ParameterName = parm.ParameterName.ToUpper
            Next
        End If


        MyBase.PrepareCommand(command, connection, transaction, commandType, commandText, commandParameters)


        If Not commandParameters Is Nothing Then
            i = 0
            For Each parm In command.Parameters

                parm.ParameterName = parm.ParameterName.Replace("@", "")

                paramName = parm.ParameterName

                'limita parametro em 30 caracteres 
                If paramName.Length > 30 Then
                    i += 1
                    paramShort = paramName.Substring(0, 25) & i
                    parm.ParameterName = paramShort
                    command.CommandText = command.CommandText.Replace(paramName, paramShort)
                End If


                Select Case parm.DbType
                    Case DbType.AnsiString, DbType.String, DbType.AnsiStringFixedLength, DbType.StringFixedLength
                        If Not IsDBNull(parm.Value) Then
                            If CStr(parm.Value) = "" Then
                                parm.Value = DBNull.Value
                            End If
                        End If
                    Case DbType.Boolean
                        'parm.DbType = DbType.Byte

                        If IsDBNull(parm.Value) Then
                            parm.Value = CByte(0)
                        Else
                            parm.Value = CByte(Math.Abs(CInt(parm.Value)))
                        End If

                End Select

            Next
        End If

        command.CommandText = command.CommandText.Replace("[", """")
        command.CommandText = command.CommandText.Replace("]", """")

        Dim objRegExParameters As New Regex("\s+AS\s+", RegexOptions.IgnoreCase)
        command.CommandText = objRegExParameters.Replace(command.CommandText, " ")

        ' Assign the replace method to the MatchEvaluator delegate.
        Dim myEvaluator As MatchEvaluator = New MatchEvaluator(AddressOf ReplaceLike)

        objRegExParameters = New Regex("\s+[^\s\(]+\s+LIKE\s+[^\s\)]+", RegexOptions.IgnoreCase)
        If objRegExParameters.IsMatch(command.CommandText) Then
            command.CommandText = objRegExParameters.Replace(command.CommandText, myEvaluator)
        End If

    End Sub

    Private Function ReplaceLike(ByVal m As Match) As String
        ' Replace each Regex match with the number of the match occurrence.
        Dim Value As String
        Dim regExLike As New Regex("\s+LIKE\s+", RegexOptions.IgnoreCase)
        'Value = regExLike.Replace(m.Value(), ",")
        'Value = Value.Replace("%", "")
        'Value = " REGEXP_LIKE(" & Value & ",'i') "
        Value = " UPPER(" & regExLike.Replace(m.Value(), ") LIKE UPPER(") & ") ESCAPE '!' "
        Return Value
    End Function


    Public Overrides Function NewConnection(Optional ByVal sConnectionString As String = "") As IDbConnection

        If sConnectionString = "" Then
            NewConnection = New ConnectionWithExtraInfo(New OracleConnection)
        Else
            NewConnection = New ConnectionWithExtraInfo(New OracleConnection(sConnectionString))
        End If

    End Function

    Private Function FormatParameterName(ByVal paramName As String) As String
        FormatParameterName = paramName 'paramName.Replace("@", "")
    End Function

    Public Overloads Overrides Function NewParameter() As IDataParameter
        NewParameter = New OracleParameter
    End Function

    Public Overloads Overrides Function NewParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType) As IDataParameter
        'Dim dataType As OracleType = CType(m_DataTypeList(dbType), OracleType)
        NewParameter = New OracleParameter
        NewParameter.DbType = dbType
        NewParameter.ParameterName = FormatParameterName(parameterName)
    End Function
    Public Overloads Overrides Function NewParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType, ByVal size As Integer) As IDataParameter
        Dim objNewParameter As New OracleParameter
        objNewParameter.DbType = dbType
        objNewParameter.ParameterName = FormatParameterName(parameterName)
        objNewParameter.Size = size
        NewParameter = objNewParameter
    End Function

    Public Overloads Overrides Function NewParameter(ByVal parameterName As String, ByVal value As Object) As IDataParameter
        NewParameter = New OracleParameter(FormatParameterName(parameterName), value)
    End Function

    Public Overloads Overrides Function NewParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType, ByVal value As Object) As IDataParameter
        Dim objNewParameter As New OracleParameter
        objNewParameter.ParameterName = FormatParameterName(parameterName)
        If dbType = dbType.DateTime Then
            objNewParameter.OracleType = OracleType.DateTime
        Else
            objNewParameter.DbType = dbType
        End If
        objNewParameter.Value = value
        NewParameter = CType(objNewParameter, IDataParameter)
    End Function



    Protected Overrides Function NewCommand() As IDbCommand
        NewCommand = New OracleCommand
    End Function

    Protected Overrides Function NewDataAdapter(ByRef cmd As IDbCommand) As IDataAdapter
        NewDataAdapter = New OracleDataAdapter(CType(cmd, OracleCommand))
    End Function

    Protected Overrides Function NewDBHelperParameterCache() As IDBHelperParameterCache
        NewDBHelperParameterCache = New MSOracleHelperParameterCache
    End Function

#Region "UpadeteDataSet"

    Private Sub ChangeEmptyStringToNull(ByRef dt As DataTable)

        Dim dr As DataRow
        Dim col As DataColumn
        Dim stringColumns As New ArrayList
        Dim i As Integer

        For Each col In dt.Columns
            If col.DataType Is GetType(String) Then
                stringColumns.Add(col.ColumnName)
            End If
        Next

        If stringColumns.Count < 0 Then
            Exit Sub
        End If
        Dim colName As String

        For Each dr In dt.Rows
            For i = 0 To stringColumns.Count - 1
                colName = CStr(stringColumns(i))
                If Not IsDBNull(dr(colName)) Then
                    If CStr(dr(colName)) = "" Then
                        dr(colName) = DBNull.Value
                    End If
                End If
            Next i
        Next
    End Sub

    Private Overloads Function UpdateDatasetTransaction(ByVal connection As IDbConnection, ByVal transaction As IDbTransaction, _
                                                      ByVal commandText As String, ByVal Table As DataTable) As Integer
        'trwo  error messages 

        Dim oDA As IDataAdapter
        Dim cdm As IDbCommand

        cdm = NewCommand()
        PrepareCommand(cdm, connection, transaction, CommandType.Text, commandText, Nothing)
        oDA = NewDataAdapter(cdm)

        Dim oCB As OracleCommandBuilder = New OracleCommandBuilder(CType(oDA, OracleDataAdapter))

        Dim rowsAffected As Integer

        oCB.QuotePrefix = Me.QuotePrefix
        oCB.QuoteSuffix = Me.QuoteSuffix

        Try
            ChangeEmptyStringToNull(Table)
            rowsAffected = CType(oDA, System.Data.Common.DbDataAdapter).Update(Table)
        Catch ex As Exception
            System.Diagnostics.Debug.WriteLine(ex.Message())
            System.Diagnostics.Debug.WriteLine(oCB.GetInsertCommand.CommandText())
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


        Dim oDA As OracleDataAdapter
        oDA = New OracleDataAdapter
        Dim cdm As IDbCommand

        'Update Command
        If UpdateCommandText.Trim() <> "" Then
            cdm = NewCommand()
            PrepareCommand(cdm, connection, transaction, UpdateCommandType, UpdateCommandText, UpdatedataParam)
            oDA.UpdateCommand = CType(cdm, OracleCommand)
        End If

        'Insert  Command 
        If InsertCommandText.Trim() <> "" Then
            cdm = NewCommand()
            PrepareCommand(cdm, connection, transaction, InsertCommandType, InsertCommandText, InsertdataParam)
            oDA.InsertCommand = CType(cdm, OracleCommand)
        End If

        'Delete  Command 
        If DeleteCommandText.Trim() <> "" Then
            cdm = NewCommand()
            PrepareCommand(cdm, connection, transaction, DeleteCommandType, DeleteCommandText, DeletedataParam)
            oDA.DeleteCommand = CType(cdm, OracleCommand)
        End If

        Dim iNumRowsAffected As Integer

        With oDA
            ChangeEmptyStringToNull(Table)
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

        Dim dbTransaction As IDbTransaction = GetConnectionTransaction(connection)

        Return UpdateDatasetTransaction(connection, dbTransaction, commandText, Table)

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

    Protected Overrides Function SQLParamReplaceString() As String
        SQLParamReplaceString = ":"
    End Function
#End Region

End Class


' SqlHelperParameterCache provides functions to leverage a static cache of procedure parameters, and the
' ability to discover parameters for stored procedures at run-time.
Public NotInheritable Class MSOracleHelperParameterCache
    Implements IDBHelperParameterCache

#Region "private methods, variables, and constructors"


    'Since this class provides only static methods, make the default constructor private to prevent 
    'instances from being created with "new SqlHelperParameterCache()".
    Friend Sub New()
    End Sub 'New 

    Private paramCache As Hashtable = Hashtable.Synchronized(New Hashtable)

    ' resolve at run time the appropriate set of OracleParameters for a stored procedure
    ' Parameters:
    ' - connectionString - a valid connection string for a SqlConnection
    ' - spName - the name of the stored procedure
    ' - includeReturnValueParameter - whether or not to include their return value parameter>
    ' Returns: SqlParameter()
    Private Function DiscoverSpParameterSet(ByVal connectionString As String, _
                                                   ByVal spName As String, _
                                                   ByVal includeReturnValueParameter As Boolean, _
                                                   ByVal ParamArray parameterValues() As Object) As IDataParameter()

        Dim cn As New OracleConnection(connectionString)
        Dim cmd As OracleCommand = New OracleCommand(spName, cn)
        Dim discoveredParameters() As IDataParameter

        Try
            cn.Open()
            cmd.CommandType = CommandType.StoredProcedure
            OracleCommandBuilder.DeriveParameters(cmd)
            If Not includeReturnValueParameter Then
                cmd.Parameters.RemoveAt(0)
            End If

            discoveredParameters = New OracleParameter(cmd.Parameters.Count - 1) {}
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

