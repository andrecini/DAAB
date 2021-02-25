Imports NUnit.Framework
Imports System.Runtime.Remoting
Imports System.Runtime.Remoting.Channels
Imports System.Diagnostics
Imports ApplicationBlocks.Data

'<TestFixture(), Explicit(), Category("Teste")> _
<TestFixture(), Category("OLEDB")> _
Public Class OleDBHelperTests
    Inherits DbHelperTests

    <TestFixtureSetUp()> _
    Public Overloads Sub SetUp()

        Dim oleDbHelper As New OleDBDbHelper

        Dim dbPath = IO.Path.Combine(CurrentPath(), "db50.mdb")
        Dim conString = "PROVIDER=Microsoft.Jet.OLEDB.4.0;Jet OLEDB:Database Password=98750; Data Source={0}"

        conString = String.Format(conString, dbPath)

        m_dbHelper = oleDbHelper
        m_dbConn = oleDbHelper.NewConnection(conString)
        m_dbConn.Open()

    End Sub

    <TestFixtureTearDown()> _
    Public Sub TearDown()
        m_dbConn.Close()
    End Sub


End Class

