Public Class ADOXUtil
    Public Shared Sub CreateEmptyAccessDatabase(ByVal fileName As String, ByVal password As String)
        Dim objCatalog As Object = Nothing
        Dim objConn As Object = Nothing

        objCatalog = Activator.CreateInstance(Type.GetTypeFromProgID("ADOX.Catalog"))

        Try
            Dim params(0) As Object
            params(0) = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                       "Data Source=" & fileName & ";" & _
                       "Jet OLEDB:Engine Type=5;Jet OLEDB:Database Password=" & password

            'invoke a Create method of a ADOX object
            'pass Parameters array
            objCatalog.GetType().InvokeMember("Create", _
                        System.Reflection.BindingFlags.InvokeMethod, _
                        Nothing, _
                        objCatalog, _
                        params)

            objConn = objCatalog.GetType().InvokeMember("ActiveConnection", _
                        System.Reflection.BindingFlags.GetProperty, _
                        Nothing, _
                         objCatalog, _
                        Nothing)

            objConn.GetType().InvokeMember("Close", _
                       System.Reflection.BindingFlags.InvokeMethod, _
                        Nothing, _
                        objConn, _
                        Nothing)

        Catch ex As Exception
            Throw ex
        Finally

            If Not (objCatalog Is Nothing) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objCatalog)
            End If

            If Not (objConn Is Nothing) Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(objConn)
            End If

            objCatalog = Nothing
            objConn = Nothing
        End Try

    End Sub

    Public Shared Function CreateADOConnection(ByVal connString As String) As Object
        Dim objConn As Object = Nothing
        objConn = Activator.CreateInstance(Type.GetTypeFromProgID("ADODB.Connection"))

        objConn.Open(connString)

        Return objConn

    End Function

    Public Shared Sub ReleaseADOConnection(ByVal cnn As Object)

        If Not (cnn Is Nothing) Then
            Try
                cnn.Close()
            Catch ex As Exception

            End Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(cnn)
        End If
    End Sub

    Public Shared Function CreateADOXCatalog() As Object
        Return Activator.CreateInstance(Type.GetTypeFromProgID("ADOX.Catalog"))
    End Function

    Public Shared Sub ReleaseADOXCatalog(ByVal cat As Object)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(cat)
    End Sub

    Public Shared Sub RenameTable(ByVal ADOconnString As String, ByVal oldTableName As String, ByVal newTableName As String)
        Dim cnnADO As Object

        Try
            cnnADO = CreateADOConnection(ADOconnString)
            RenameTable(cnnADO, oldTableName, newTableName)
        Finally
            ADOXUtil.ReleaseADOConnection(cnnADO)
        End Try

    End Sub


    Public Shared Sub RenameTable(ByVal connADO As Object, ByVal oldTableName As String, ByVal newTableName As String)
        Dim objCatalog As Object = Nothing

        Dim i As Integer
        Dim bFound As Boolean

        Try
            objCatalog = CreateADOXCatalog()
            objCatalog.ActiveConnection = connADO

            'Change the name...
            For i = 0 To objCatalog.Tables.Count - 1
                If objCatalog.Tables(i).Name = oldTableName Then
                    bFound = True
                    Exit For
                End If
            Next i

            If bFound Then
                objCatalog.Tables(oldTableName).Name = newTableName
            End If

        Finally
            ReleaseADOXCatalog(objCatalog)
        End Try

    End Sub


    Public Shared Sub RenameTableColumn(ByVal ADOconnString As String, ByVal tableName As String, ByVal oldColummName As String, ByVal newColummName As String)

        Dim cnnADO As Object

        Try
            cnnADO = CreateADOConnection(ADOconnString)
            RenameTableColumn(cnnADO, tableName, oldColummName, newColummName)
        Finally
            ADOXUtil.ReleaseADOConnection(cnnADO)
        End Try

    End Sub

    Public Shared Sub RenameTableColumn(ByVal connADO As Object, ByVal tableName As String, ByVal oldColummName As String, ByVal newColummName As String)
        Dim objCatalog As Object = Nothing
        Dim bFound As Boolean
        Dim i As Integer

        Try
            objCatalog = CreateADOXCatalog()
            objCatalog.ActiveConnection = connADO

            For i = 0 To objCatalog.Tables(tableName).Columns.count - 1
                If objCatalog.Tables(tableName).Columns(i).Name = oldColummName Then
                    bFound = True
                    Exit For
                End If
            Next i

            If bFound Then
                objCatalog.Tables(tableName).Columns(oldColummName).Name = newColummName
                objCatalog.Tables.Refresh()
            End If

        Finally
            ReleaseADOXCatalog(objCatalog)
        End Try

    End Sub



    Public Shared Sub AllowZeroLength(ByVal ADOconnString As String, ByVal tableName As String, ByVal ColummName As String)
        Dim cnnADO As Object

        Try
            cnnADO = CreateADOConnection(ADOconnString)
            AllowZeroLength(cnnADO, tableName, ColummName)
        Finally
            ADOXUtil.ReleaseADOConnection(cnnADO)
        End Try
    End Sub


    Public Shared Sub AllowZeroLength(ByVal connADO As Object, ByVal tableName As String, ByVal ColummName As String)
        Dim objCatalog As Object = Nothing
        Dim objColumn As Object
        Dim bFound As Boolean
        Dim i As Integer

        Try
            objCatalog = CreateADOXCatalog()
            objCatalog.ActiveConnection = connADO

            objColumn = objCatalog.Tables(tableName).Columns(ColummName)
            objColumn.Properties("Jet OLEDB:Allow Zero Length") = True

        Finally
            ReleaseADOXCatalog(objCatalog)
        End Try

    End Sub


    Public Shared Function GetDatabaseTables(ByVal ADOconnString As String) As String()
        Dim cnnADO As Object
        Dim tables() As String

        Try
            ReDim tables(-1)
            cnnADO = CreateADOConnection(ADOconnString)
            tables = GetDatabaseTables(cnnADO)
        Finally
            ADOXUtil.ReleaseADOConnection(cnnADO)
        End Try

        Return tables

    End Function

    Public Shared Function GetDatabaseTables(ByVal connADO As Object) As String()

        Dim objCatalog As Object = Nothing
        objCatalog = CreateADOXCatalog()
        objCatalog.ActiveConnection = connADO
        Dim i As Integer
        Dim tables() As String
        Dim tableList As New ArrayList


        For i = 0 To objCatalog.Tables.Count - 1

            If objCatalog.Tables(i).Type = "TABLE" Then
                tableList.Add(objCatalog.Tables(i).Name)
            End If
        Next

        ReleaseADOXCatalog(objCatalog)

        ReDim tables(tableList.Count - 1)

        tableList.CopyTo(tables)

        Return tables

    End Function


End Class
