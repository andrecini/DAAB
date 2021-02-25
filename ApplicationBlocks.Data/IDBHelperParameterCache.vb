Public Interface IDBHelperParameterCache

#Region "caching functions"
    ' add parameter array to the cache
    ' Parameters
    ' -connectionString - a valid connection string for a SqlConnection 
    ' -commandText - the stored procedure name or T-SQL command 
    ' -commandParameters - an array of SqlParamters to be cached 
    Sub CacheParameterSet(ByVal connectionString As String, _
                                        ByVal commandText As String, _
                                        ByVal ParamArray commandParameters() As IDataParameter)

    ' retrieve a parameter array from the cache
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection 
    ' -commandText - the stored procedure name or T-SQL command 
    ' Returns: an array of SqlParamters 
    Function GetCachedParameterSet(ByVal connectionString As String, ByVal commandText As String) As IDataParameter()

#End Region

#Region "Parameter Discovery Functions"

    ' Retrieves the set of SqlParameters appropriate for the stored procedure
    ' 
    ' This method will query the database for this information, and then store it in a cache for future requests.
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection 
    ' -spName - the name of the stored procedure 
    ' Returns: an array of SqlParameters
    Overloads Function GetSpParameterSet(ByVal connectionString As String, ByVal spName As String) As IDataParameter()


    ' Retrieves the set of SqlParameters appropriate for the stored procedure
    ' 
    ' This method will query the database for this information, and then store it in a cache for future requests.
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection
    ' -spName - the name of the stored procedure 
    ' -includeReturnValueParameter - a bool value indicating whether the return value parameter should be included in the results 
    ' Returns: an array of SqlParameters 
    Overloads Function GetSpParameterSet(ByVal connectionString As String, _
                                                       ByVal spName As String, _
                                                       ByVal includeReturnValueParameter As Boolean) As IDataParameter()

#End Region

End Interface
