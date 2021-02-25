Option Explicit On 
Option Strict On

' The IDBHelper interface 

Public Interface IDBHelper

#Region "Parameter Cache"
    ReadOnly Property DBHelperParameterCache() As IDBHelperParameterCache
#End Region

#Region "Null Functions"
    Overloads Function Null2Number(ByVal value As Object) As Object
    Overloads Function Null2String(ByVal value As Object) As Object
#End Region

#Region "Create Database Objects"
    Overloads Function NewConnection(Optional ByVal sConnectionString As String = "") As IDbConnection
    Overloads Function NewParameter() As IDataParameter
    Overloads Function NewParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType) As IDataParameter
    Overloads Function NewParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType, ByVal size As Integer) As IDataParameter
    Overloads Function NewParameter(ByVal parameterName As String, ByVal value As Object) As IDataParameter
    Overloads Function NewParameter(ByVal parameterName As String, ByVal dbType As System.Data.DbType, ByVal value As Object) As IDataParameter
    Overloads Function GetDbType(ByVal [type] As Type) As DbType
#End Region

#Region "ExecuteNonQuery"
    ' Execute a SqlCommand (that returns no resultset and takes no parameters) against the database specified in 
    ' the connection string. 
    ' e.g.:  
    '  Dim result as Integer =  ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders")
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection
    ' -commandType - the CommandType (stored procedure, text, etc.)
    ' -commandText - the stored procedure name or T- command
    ' Returns: an int representing the number of rows affected by the command
    Overloads Function ExecuteNonQuery(ByVal connectionString As String, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String) As Integer


    ' Execute a Command (that returns no resultset) against the database specified in the connection string 
    ' using the provided parameters.
    ' e.g.:  
    ' Dim result as Integer = ExecuteNonQuery(connString, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24))
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection
    ' -commandType - the CommandType (stored procedure, text, etc.)
    ' -commandText - the stored procedure name or T-SQL command
    ' -commandParameters - an array of SqlParamters used to execute the command
    ' Returns: an int representing the number of rows affected by the command
    Overloads Function ExecuteNonQuery(ByVal connectionString As String, _
                                                     ByVal commandType As CommandType, _
                                                     ByVal commandText As String, _
                                                     ByVal ParamArray commandParameters() As IDataParameter) As Integer

    ' Execute a stored procedure via a SqlCommand (that returns no resultset) against the database specified in 
    ' the connection string using the provided parameter values.  This method will discover the parameters for the 
    ' stored procedure, and assign the values based on parameter order.
    ' This method provides no access to output parameters or the stored procedure's return value parameter.
    ' e.g.:  
    '  Dim result as Integer = ExecuteNonQuery(connString, "PublishOrders", 24, 36)
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection
    ' -spName - the name of the stored procedure
    ' -parameterValues - an array of objects to be assigned as the input values of the stored procedure
    ' Returns: an int representing the number of rows affected by the command
    Overloads Function ExecuteNonQuery(ByVal connectionString As String, _
                                                     ByVal spName As String, _
                                                     ByVal ParamArray parameterValues() As Object) As Integer

    ' Execute a SqlCommand (that returns no resultset and takes no parameters) against the provided SqlConnection. 
    ' e.g.:  
    ' Dim result as Integer = ExecuteNonQuery(conn, CommandType.StoredProcedure, "PublishOrders")
    ' Parameters:
    ' -connection - a valid SqlConnection
    ' -commandType - the CommandType (stored procedure, text, etc.)
    ' -commandText - the stored procedure name or T-SQL command 
    ' Returns: an int representing the number of rows affected by the command
    Overloads Function ExecuteNonQuery(ByVal connection As IDbConnection, _
                                                     ByVal commandType As CommandType, _
                                                     ByVal commandText As String) As Integer


    ' Execute a SqlCommand (that returns no resultset) against the specified SqlConnection 
    ' using the provided parameters.
    ' e.g.:  
    '  Dim result as Integer = ExecuteNonQuery(conn, CommandType.StoredProcedure, "PublishOrders", new SqlParameter("@prodid", 24))
    ' Parameters:
    ' -connection - a valid SqlConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.)
    ' -commandText - the stored procedure name or T-SQL command 
    ' -commandParameters - an array of SqlParamters used to execute the command 
    ' Returns: an int representing the number of rows affected by the command 
    Overloads Function ExecuteNonQuery(ByVal connection As IDbConnection, _
                                                     ByVal commandType As CommandType, _
                                                     ByVal commandText As String, _
                                                     ByVal ParamArray commandParameters() As IDataParameter) As Integer


    ' Execute a stored procedure via a SqlCommand (that returns no resultset) against the specified SqlConnection 
    ' using the provided parameter values.  This method will discover the parameters for the 
    ' stored procedure, and assign the values based on parameter order.
    ' This method provides no access to output parameters or the stored procedure's return value parameter.
    ' e.g.:  
    '  Dim result as integer = ExecuteNonQuery(conn, "PublishOrders", 24, 36)
    ' Parameters:
    ' -connection - a valid SqlConnection
    ' -spName - the name of the stored procedure 
    ' -parameterValues - an array of objects to be assigned as the input values of the stored procedure 
    ' Returns: an int representing the number of rows affected by the command 
    Overloads Function ExecuteNonQuery(ByVal connection As IDbConnection, _
                                                     ByVal spName As String, _
                                                     ByVal ParamArray parameterValues() As Object) As Integer



    ' Execute a SqlCommand (that returns no resultset and takes no parameters) against the provided SqlTransaction.
    ' e.g.:  
    '  Dim result as Integer = ExecuteNonQuery(trans, CommandType.StoredProcedure, "PublishOrders")
    ' Parameters:
    ' -transaction - a valid SqlTransaction associated with the connection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' Returns: an int representing the number of rows affected by the command 
    Overloads Function ExecuteNonQuery(ByVal transaction As IDbTransaction, _
                                                     ByVal commandType As CommandType, _
                                                     ByVal commandText As String) As Integer


    ' Execute a SqlCommand (that returns no resultset) against the specified SqlTransaction
    ' using the provided parameters.
    ' e.g.:  
    ' Dim result as Integer = ExecuteNonQuery(trans, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))
    ' Parameters:
    ' -transaction - a valid SqlTransaction 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' -commandParameters - an array of SqlParamters used to execute the command 
    ' Returns: an int representing the number of rows affected by the command 
    Overloads Function ExecuteNonQuery(ByVal transaction As IDbTransaction, _
                                                     ByVal commandType As CommandType, _
                                                     ByVal commandText As String, _
                                                     ByVal ParamArray commandParameters() As IDataParameter) As Integer

    ' Execute a stored procedure via a SqlCommand (that returns no resultset) against the specified SqlTransaction 
    ' using the provided parameter values.  This method will discover the parameters for the 
    ' stored procedure, and assign the values based on parameter order.
    ' This method provides no access to output parameters or the stored procedure's return value parameter.
    ' e.g.:  
    ' Dim result As Integer = SqlHelper.ExecuteNonQuery(trans, "PublishOrders", 24, 36)
    ' Parameters:
    ' -transaction - a valid SqlTransaction 
    ' -spName - the name of the stored procedure 
    ' -parameterValues - an array of objects to be assigned as the input values of the stored procedure 
    ' Returns: an int representing the number of rows affected by the command 
    Overloads Function ExecuteNonQuery(ByVal transaction As IDbTransaction, _
                                                     ByVal spName As String, _
                                                     ByVal ParamArray parameterValues() As Object) As Integer



#End Region

#Region "ExecuteDataset"

    ' Execute a SqlCommand (that returns a resultset and takes no parameters) against the database specified in 
    ' the connection string. 
    ' e.g.:  
    ' Dim ds As DataSet = SqlHelper.ExecuteDataset("", commandType.StoredProcedure, "GetOrders")
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection
    ' -commandType - the CommandType (stored procedure, text, etc.)
    ' -commandText - the stored procedure name or T-SQL command
    ' Returns: a dataset containing the resultset generated by the command
    Overloads Function ExecuteDataset(ByVal connectionString As String, _
                                                    ByVal commandType As CommandType, _
                                                    ByVal commandText As String) As DataSet

    ' Execute a SqlCommand (that returns a resultset) against the database specified in the connection string 
    ' using the provided parameters.
    ' e.g.:  
    ' Dim ds as Dataset = ExecuteDataset(connString, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection
    ' -commandType - the CommandType (stored procedure, text, etc.)
    ' -commandText - the stored procedure name or T-SQL command
    ' -commandParameters - an array of SqlParamters used to execute the command
    ' Returns: a dataset containing the resultset generated by the command
    Overloads Function ExecuteDataset(ByVal connectionString As String, _
                                                    ByVal commandType As CommandType, _
                                                    ByVal commandText As String, _
                                                    ByVal ParamArray commandParameters() As IDataParameter) As DataSet



    ' Execute a stored procedure via a SqlCommand (that returns a resultset) against the database specified in 
    ' the connection string using the provided parameter values.  This method will discover the parameters for the 
    ' stored procedure, and assign the values based on parameter order.
    ' This method provides no access to output parameters or the stored procedure's return value parameter.
    ' e.g.:  
    ' Dim ds as Dataset= ExecuteDataset(connString, "GetOrders", 24, 36)
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection
    ' -spName - the name of the stored procedure
    ' -parameterValues - an array of objects to be assigned as the input values of the stored procedure
    ' Returns: a dataset containing the resultset generated by the command
    Overloads Function ExecuteDataset(ByVal connectionString As String, _
                                                    ByVal spName As String, _
                                                    ByVal ParamArray parameterValues() As Object) As DataSet

    ' Execute a SqlCommand (that returns a resultset and takes no parameters) against the provided SqlConnection. 
    ' e.g.:  
    ' Dim ds as Dataset = ExecuteDataset(conn, CommandType.StoredProcedure, "GetOrders")
    ' Parameters:
    ' -connection - a valid SqlConnection
    ' -commandType - the CommandType (stored procedure, text, etc.)
    ' -commandText - the stored procedure name or T-SQL command
    ' Returns: a dataset containing the resultset generated by the command
    Overloads Function ExecuteDataset(ByVal connection As IDbConnection, _
                                                    ByVal commandType As CommandType, _
                                                    ByVal commandText As String) As DataSet


    ' Execute a SqlCommand (that returns a resultset) against the specified SqlConnection 
    ' using the provided parameters.
    ' e.g.:  
    ' Dim ds as Dataset = ExecuteDataset(conn, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))
    ' Parameters:
    ' -connection - a valid SqlConnection
    ' -commandType - the CommandType (stored procedure, text, etc.)
    ' -commandText - the stored procedure name or T-SQL command
    ' -commandParameters - an array of SqlParamters used to execute the command
    ' Returns: a dataset containing the resultset generated by the command
    Overloads Function ExecuteDataset(ByVal connection As IDbConnection, _
                                                    ByVal commandType As CommandType, _
                                                    ByVal commandText As String, _
                                                    ByVal ParamArray commandParameters() As IDataParameter) As DataSet

    ' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified SqlConnection 
    ' using the provided parameter values.  This method will discover the parameters for the 
    ' stored procedure, and assign the values based on parameter order.
    ' This method provides no access to output parameters or the stored procedure's return value parameter.
    ' e.g.:  
    ' Dim ds As Dataset = ExecuteDataset(conn, "GetOrders", 24, 36)
    ' Parameters:
    ' -connection - a valid SqlConnection
    ' -spName - the name of the stored procedure
    ' -parameterValues - an array of objects to be assigned as the input values of the stored procedure
    ' Returns: a dataset containing the resultset generated by the command
    Overloads Function ExecuteDataset(ByVal connection As IDbConnection, _
                                                    ByVal spName As String, _
                                                    ByVal ParamArray parameterValues() As Object) As DataSet




    ' Execute a SqlCommand (that returns a resultset and takes no parameters) against the provided SqlTransaction. 
    ' e.g.:  
    ' Dim ds As Dataset = ExecuteDataset(trans, CommandType.StoredProcedure, "GetOrders")
    ' Parameters
    ' -transaction - a valid SqlTransaction
    ' -commandType - the CommandType (stored procedure, text, etc.)
    ' -commandText - the stored procedure name or T-SQL command
    ' Returns: a dataset containing the resultset generated by the command
    Overloads Function ExecuteDataset(ByVal transaction As IDbTransaction, _
                                                    ByVal commandType As CommandType, _
                                                    ByVal commandText As String) As DataSet



    ' Execute a SqlCommand (that returns a resultset) against the specified SqlTransaction
    ' using the provided parameters.
    ' e.g.:  
    ' Dim ds As Dataset = ExecuteDataset(trans, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))
    ' Parameters
    ' -transaction - a valid SqlTransaction 
    ' -commandType - the CommandType (stored procedure, text, etc.)
    ' -commandText - the stored procedure name or T-SQL command
    ' -commandParameters - an array of SqlParamters used to execute the command
    ' Returns: a dataset containing the resultset generated by the command
    Overloads Function ExecuteDataset(ByVal transaction As IDbTransaction, _
                                                    ByVal commandType As CommandType, _
                                                    ByVal commandText As String, _
                                                    ByVal ParamArray commandParameters() As IDataParameter) As DataSet


    ' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified
    ' SqlTransaction using the provided parameter values.  This method will discover the parameters for the 
    ' stored procedure, and assign the values based on parameter order.
    ' This method provides no access to output parameters or the stored procedure's return value parameter.
    ' e.g.:  
    ' Dim ds As Dataset = ExecuteDataset(trans, "GetOrders", 24, 36)
    ' Parameters:
    ' -transaction - a valid SqlTransaction 
    ' -spName - the name of the stored procedure
    ' -parameterValues - an array of objects to be assigned as the input values of the stored procedure
    ' Returns: a dataset containing the resultset generated by the command
    Overloads Function ExecuteDataset(ByVal transaction As IDbTransaction, _
                                                    ByVal spName As String, _
                                                    ByVal ParamArray parameterValues() As Object) As DataSet


#End Region

#Region "ExecuteReader"

    ' Execute a SqlCommand (that returns a resultset and takes no parameters) against the database specified in 
    ' the connection string. 
    ' e.g.:  
    ' Dim dr As SqlDataReader = ExecuteReader(connString, CommandType.StoredProcedure, "GetOrders")
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' Returns: a SqlDataReader containing the resultset generated by the command 
    Overloads Function ExecuteReader(ByVal connectionString As String, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String) As IDataReader

    ' Execute a SqlCommand (that returns a resultset) against the database specified in the connection string 
    ' using the provided parameters.
    ' e.g.:  
    ' Dim dr As SqlDataReader = ExecuteReader(connString, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' -commandParameters - an array of SqlParamters used to execute the command 
    ' Returns: a SqlDataReader containing the resultset generated by the command 
    Overloads Function ExecuteReader(ByVal connectionString As String, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String, _
                                                   ByVal ParamArray commandParameters() As IDataParameter) As IDataReader

    ' Execute a stored procedure via a SqlCommand (that returns a resultset) against the database specified in 
    ' the connection string using the provided parameter values.  This method will discover the parameters for the 
    ' stored procedure, and assign the values based on parameter order.
    ' This method provides no access to output parameters or the stored procedure's return value parameter.
    ' e.g.:  
    ' Dim dr As SqlDataReader = ExecuteReader(connString, "GetOrders", 24, 36)
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection 
    ' -spName - the name of the stored procedure 
    ' -parameterValues - an array of objects to be assigned as the input values of the stored procedure 
    ' Returns: a SqlDataReader containing the resultset generated by the command 
    Overloads Function ExecuteReader(ByVal connectionString As String, _
                                                   ByVal spName As String, _
                                                   ByVal ParamArray parameterValues() As Object) As IDataReader

    ' Execute a SqlCommand (that returns a resultset and takes no parameters) against the provided SqlConnection. 
    ' e.g.:  
    ' Dim dr As SqlDataReader = ExecuteReader(conn, CommandType.StoredProcedure, "GetOrders")
    ' Parameters:
    ' -connection - a valid SqlConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' Returns: a SqlDataReader containing the resultset generated by the command 
    Overloads Function ExecuteReader(ByVal connection As IDbConnection, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String) As IDataReader



    ' Execute a SqlCommand (that returns a resultset) against the specified SqlConnection 
    ' using the provided parameters.
    ' e.g.:  
    ' Dim dr As SqlDataReader = ExecuteReader(conn, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))
    ' Parameters:
    ' -connection - a valid SqlConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' -commandParameters - an array of SqlParamters used to execute the command 
    ' Returns: a SqlDataReader containing the resultset generated by the command 
    Overloads Function ExecuteReader(ByVal connection As IDbConnection, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String, _
                                                   ByVal ParamArray commandParameters() As IDataParameter) As IDataReader


    ' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified SqlConnection 
    ' using the provided parameter values.  This method will discover the parameters for the 
    ' stored procedure, and assign the values based on parameter order.
    ' This method provides no access to output parameters or the stored procedure's return value parameter.
    ' e.g.:  
    ' Dim dr As SqlDataReader = ExecuteReader(conn, "GetOrders", 24, 36)
    ' Parameters:
    ' -connection - a valid SqlConnection 
    ' -spName - the name of the stored procedure 
    ' -parameterValues - an array of objects to be assigned as the input values of the stored procedure 
    ' Returns: a SqlDataReader containing the resultset generated by the command 
    Overloads Function ExecuteReader(ByVal connection As IDbConnection, _
                                                   ByVal spName As String, _
                                                   ByVal ParamArray parameterValues() As Object) As IDataReader

    ' Execute a SqlCommand (that returns a resultset and takes no parameters) against the provided SqlTransaction.
    ' e.g.:  
    ' Dim dr As SqlDataReader = ExecuteReader(trans, CommandType.StoredProcedure, "GetOrders")
    ' Parameters:
    ' -transaction - a valid SqlTransaction  
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' Returns: a SqlDataReader containing the resultset generated by the command 
    Overloads Function ExecuteReader(ByVal transaction As IDbTransaction, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String) As IDataReader

    ' Execute a SqlCommand (that returns a resultset) against the specified SqlTransaction
    ' using the provided parameters.
    ' e.g.:  
    ' Dim dr As SqlDataReader = ExecuteReader(trans, CommandType.StoredProcedure, "GetOrders", new SqlParameter("@prodid", 24))
    ' Parameters:
    ' -transaction - a valid SqlTransaction 
    ' -commandType - the CommandType (stored procedure, text, etc.)
    ' -commandText - the stored procedure name or T-SQL command 
    ' -commandParameters - an array of SqlParamters used to execute the command 
    ' Returns: a SqlDataReader containing the resultset generated by the command 
    Overloads Function ExecuteReader(ByVal transaction As IDbTransaction, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String, _
                                                   ByVal ParamArray commandParameters() As IDataParameter) As IDataReader


    ' Execute a stored procedure via a SqlCommand (that returns a resultset) against the specified SqlTransaction 
    ' using the provided parameter values.  This method will discover the parameters for the 
    ' stored procedure, and assign the values based on parameter order.
    ' This method provides no access to output parameters or the stored procedure's return value parameter.
    ' e.g.:  
    ' Dim dr As SqlDataReader = ExecuteReader(trans, "GetOrders", 24, 36)
    ' Parameters:
    ' -transaction - a valid SqlTransaction 
    ' -spName - the name of the stored procedure 
    ' -parameterValues - an array of objects to be assigned as the input values of the stored procedure
    ' Returns: a SqlDataReader containing the resultset generated by the command
    Overloads Function ExecuteReader(ByVal transaction As IDbTransaction, _
                                                   ByVal spName As String, _
                                                   ByVal ParamArray parameterValues() As Object) As IDataReader


#End Region

#Region "ExecuteScalar"

    ' Execute a SqlCommand (that returns a 1x1 resultset and takes no parameters) against the database specified in 
    ' the connection string. 
    ' e.g.:  
    ' Dim orderCount As Integer = CInt(ExecuteScalar(connString, CommandType.StoredProcedure, "GetOrderCount"))
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' Returns: an object containing the value in the 1x1 resultset generated by the command
    Overloads Function ExecuteScalar(ByVal connectionString As String, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String) As Object

    ' Execute a SqlCommand (that returns a 1x1 resultset) against the database specified in the connection string 
    ' using the provided parameters.
    ' e.g.:  
    ' Dim orderCount As Integer = Cint(ExecuteScalar(connString, CommandType.StoredProcedure, "GetOrderCount", new SqlParameter("@prodid", 24)))
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' -commandParameters - an array of SqlParamters used to execute the command 
    ' Returns: an object containing the value in the 1x1 resultset generated by the command 
    Overloads Function ExecuteScalar(ByVal connectionString As String, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String, _
                                                   ByVal ParamArray commandParameters() As IDataParameter) As Object

    ' Execute a stored procedure via a SqlCommand (that returns a 1x1 resultset) against the database specified in 
    ' the connection string using the provided parameter values.  This method will discover the parameters for the 
    ' stored procedure, and assign the values based on parameter order.
    ' This method provides no access to output parameters or the stored procedure's return value parameter.
    ' e.g.:  
    ' Dim orderCount As Integer = CInt(ExecuteScalar(connString, "GetOrderCount", 24, 36))
    ' Parameters:
    ' -connectionString - a valid connection string for a SqlConnection 
    ' -spName - the name of the stored procedure 
    ' -parameterValues - an array of objects to be assigned as the input values of the stored procedure 
    ' Returns: an object containing the value in the 1x1 resultset generated by the command 
    Overloads Function ExecuteScalar(ByVal connectionString As String, _
                                                  ByVal spName As String, _
                                                  ByVal ParamArray parameterValues() As Object) As Object

    ' Execute a SqlCommand (that returns a 1x1 resultset and takes no parameters) against the provided SqlConnection. 
    ' e.g.:  
    ' Dim orderCount As Integer = CInt(ExecuteScalar(conn, CommandType.StoredProcedure, "GetOrderCount"))
    ' Parameters:
    ' -connection - a valid SqlConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' Returns: an object containing the value in the 1x1 resultset generated by the command 
    Overloads Function ExecuteScalar(ByVal connection As IDbConnection, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String) As Object

    ' Execute a SqlCommand (that returns a 1x1 resultset) against the specified SqlConnection 
    ' using the provided parameters.
    ' e.g.:  
    ' Dim orderCount As Integer = CInt(ExecuteScalar(conn, CommandType.StoredProcedure, "GetOrderCount", new SqlParameter("@prodid", 24)))
    ' Parameters:
    ' -connection - a valid SqlConnection 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' -commandParameters - an array of SqlParamters used to execute the command 
    ' Returns: an object containing the value in the 1x1 resultset generated by the command 
    Overloads Function ExecuteScalar(ByVal connection As IDbConnection, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String, _
                                                   ByVal ParamArray commandParameters() As IDataParameter) As Object


    ' Execute a stored procedure via a SqlCommand (that returns a 1x1 resultset) against the specified SqlConnection 
    ' using the provided parameter values.  This method will discover the parameters for the 
    ' stored procedure, and assign the values based on parameter order.
    ' This method provides no access to output parameters or the stored procedure's return value parameter.
    ' e.g.:  
    ' Dim orderCount As Integer = CInt(ExecuteScalar(conn, "GetOrderCount", 24, 36))
    ' Parameters:
    ' -connection - a valid SqlConnection 
    ' -spName - the name of the stored procedure 
    ' -parameterValues - an array of objects to be assigned as the input values of the stored procedure 
    ' Returns: an object containing the value in the 1x1 resultset generated by the command 
    Overloads Function ExecuteScalar(ByVal connection As IDbConnection, _
                                                   ByVal spName As String, _
                                                   ByVal ParamArray parameterValues() As Object) As Object



    ' Execute a SqlCommand (that returns a 1x1 resultset and takes no parameters) against the provided SqlTransaction.
    ' e.g.:  
    ' Dim orderCount As Integer  = CInt(ExecuteScalar(trans, CommandType.StoredProcedure, "GetOrderCount"))
    ' Parameters:
    ' -transaction - a valid SqlTransaction 
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' Returns: an object containing the value in the 1x1 resultset generated by the command 
    Overloads Function ExecuteScalar(ByVal transaction As IDbTransaction, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String) As Object

    ' Execute a SqlCommand (that returns a 1x1 resultset) against the specified SqlTransaction
    ' using the provided parameters.
    ' e.g.:  
    ' Dim orderCount As Integer = CInt(ExecuteScalar(trans, CommandType.StoredProcedure, "GetOrderCount", new SqlParameter("@prodid", 24)))
    ' Parameters:
    ' -transaction - a valid SqlTransaction  
    ' -commandType - the CommandType (stored procedure, text, etc.) 
    ' -commandText - the stored procedure name or T-SQL command 
    ' -commandParameters - an array of SqlParamters used to execute the command 
    ' Returns: an object containing the value in the 1x1 resultset generated by the command 
    Overloads Function ExecuteScalar(ByVal transaction As IDbTransaction, _
                                                   ByVal commandType As CommandType, _
                                                   ByVal commandText As String, _
                                                   ByVal ParamArray commandParameters() As IDataParameter) As Object


    ' Execute a stored procedure via a SqlCommand (that returns a 1x1 resultset) against the specified SqlTransaction 
    ' using the provided parameter values.  This method will discover the parameters for the 
    ' stored procedure, and assign the values based on parameter order.
    ' This method provides no access to output parameters or the stored procedure's return value parameter.
    ' e.g.:  
    ' Dim orderCount As Integer = CInt(ExecuteScalar(trans, "GetOrderCount", 24, 36))
    ' Parameters:
    ' -transaction - a valid SqlTransaction 
    ' -spName - the name of the stored procedure 
    ' -parameterValues - an array of objects to be assigned as the input values of the stored procedure 
    ' Returns: an object containing the value in the 1x1 resultset generated by the command 
    Overloads Function ExecuteScalar(ByVal transaction As IDbTransaction, _
                                                   ByVal spName As String, _
                                                   ByVal ParamArray parameterValues() As Object) As Object

#End Region

#Region "UpdateDataset"

    Overloads Function UpdateDataset(ByVal connection As IDbConnection, _
                                                        ByVal commandText As String, ByVal Table As DataTable) As Integer

    Overloads Function UpdateDataset(ByVal connection As IDbConnection, _
                                                    ByVal commandText As String, ByVal ds As DataSet, ByVal srcTable As String) As Integer

    Overloads Function UpdateDataset(ByVal transcation As IDbTransaction, _
                                                    ByVal commandText As String, ByVal Table As DataTable) As Integer

    Overloads Function UpdateDataset(ByVal transcation As IDbTransaction, _
                                                   ByVal commandText As String, ByVal ds As DataSet, ByVal srcTable As String) As Integer

    Overloads Function UpdateDataset(ByVal connection As IDbConnection, _
                                       ByVal Table As DataTable, _
                                       Optional ByVal UpdateCommandText As String = "", _
                                       Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                       Optional ByVal InsertCommandText As String = "", _
                                       Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                       Optional ByVal DeleteCommandText As String = "", _
                                       Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer

    Overloads Function UpdateDataset(ByVal connection As IDbConnection, _
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

    Overloads Function UpdateDataset(ByVal connection As IDbConnection, _
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
                                     Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer


    Overloads Function UpdateDataset(ByVal connection As IDbConnection, _
                                       ByVal ds As DataSet, _
                                       ByVal srcTable As String, _
                                       Optional ByVal UpdateCommandText As String = "", _
                                       Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                       Optional ByVal InsertCommandText As String = "", _
                                       Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                       Optional ByVal DeleteCommandText As String = "", _
                                       Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer


    Overloads Function UpdateDataset(ByVal transcation As IDbTransaction, _
                                       ByVal Table As DataTable, _
                                       Optional ByVal UpdateCommandText As String = "", _
                                       Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                       Optional ByVal InsertCommandText As String = "", _
                                       Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                       Optional ByVal DeleteCommandText As String = "", _
                                       Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer

    Overloads Function UpdateDataset(ByVal transaction As IDbTransaction, _
                                           ByVal ds As DataSet, _
                                           ByVal srcTable As String, _
                                           Optional ByVal UpdateCommandText As String = "", _
                                           Optional ByVal UpdatedataParam() As IDataParameter = Nothing, _
                                           Optional ByVal InsertCommandText As String = "", _
                                           Optional ByVal InsertdataParam() As IDataParameter = Nothing, _
                                           Optional ByVal DeleteCommandText As String = "", _
                                           Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer

    Overloads Function UpdateDataset(ByVal transaction As IDbTransaction, _
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

    Overloads Function UpdateDataset(ByVal transaction As IDbTransaction, _
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
                                 Optional ByVal DeletedataParam() As IDataParameter = Nothing) As Integer

#End Region


End Interface
