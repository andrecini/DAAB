Imports NUnit.Framework
Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Diagnostics
Imports ApplicationBlocks.Data

'<TestFixture(), Explicit(), Category("Teste")> _
<TestFixture(), Category("Oracle")> _
Public Class OracleDBHelperTests
    Inherits DbHelperTests

    <TestFixtureSetUp()> _
     Public Overloads Sub SetUp()

        Dim oracleDbHelper As New MSOracleDbHelper
        oracleDbHelper.QuotePrefix = """"
        oracleDbHelper.QuoteSuffix = """"

        Dim conString = "Data Source=(DESCRIPTION=(ADDRESS_LIST =(ADDRESS=(PROTOCOL=TCP) (HOST=dbricardo) (PORT=1521))) (CONNECT_DATA=(SERVICE_NAME=XE))); User Id=isoplan; Password=admin;"

        oracleDbHelper.SetOracleDllDirectory("C:\oraclexe\instantclient_10_2")

        m_dbHelper = oracleDbHelper
        m_dbConn = oracleDbHelper.NewConnection(conString)
        m_dbConn.Open()

    End Sub

    <TestFixtureTearDown()> _
    Public Sub TearDown()
        m_dbConn.Close()
    End Sub

End Class