Imports NUnit.Framework
Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Diagnostics
Imports ApplicationBlocks.Data

'<TestFixture(), Explicit(), Category("Teste")> _
<TestFixture(), Category("SQL Server")> _
Public Class SQLDBHelperTests
    Inherits DbHelperTests


    <TestFixtureSetUp()> _
    Public Overloads Sub SetUp()

        Dim sqlDbHelper As New SqlDbHelper
        sqlDbHelper.QuotePrefix = "["
        sqlDbHelper.QuoteSuffix = "]"

        Dim conString = "Data Source=DBRICARDO\SQLEXPRESS;Initial Catalog=isoplan;User ID=sa;Password=admin;"

        m_dbHelper = sqlDbHelper
        m_dbConn = sqlDbHelper.NewConnection(conString)
        m_dbConn.Open()

    End Sub

    <TestFixtureTearDown()> _
    Public Sub TearDown()
        m_dbConn.Close()
    End Sub

End Class