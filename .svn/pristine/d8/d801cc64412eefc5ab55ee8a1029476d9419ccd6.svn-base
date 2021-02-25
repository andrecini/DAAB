
Imports System.Data.Common


Public Class NestedDbTransaction
    Inherits DbTransaction
    Implements IDisposable



    Private ReadOnly m_dbTransaction As DbTransaction
    Private ReadOnly m_nestedTransaction As Boolean

    Public ReadOnly Property Transaction() As IDbTransaction
        Get
            Return m_dbTransaction
        End Get
    End Property

    Sub New(ByVal dbTransaction As DbTransaction, ByVal nestedTransaction As Boolean)
        m_dbTransaction = dbTransaction
        m_nestedTransaction = nestedTransaction
    End Sub

    Public Overrides Sub Commit()
        If Not m_nestedTransaction Then
            m_dbTransaction.Commit()
        End If
    End Sub

    Protected Overrides ReadOnly Property DbConnection() As System.Data.Common.DbConnection
        Get
            Return m_dbTransaction.Connection
        End Get
    End Property

    Public Overrides ReadOnly Property IsolationLevel() As System.Data.IsolationLevel
        Get
            Return m_dbTransaction.IsolationLevel
        End Get
    End Property

    Public Overrides Sub Rollback()

        If Not m_dbTransaction.Connection Is Nothing Then
            m_dbTransaction.Rollback()
        End If

    End Sub

    Public Sub Dispose() Implements System.IDisposable.Dispose
        If Not m_nestedTransaction Then
            MyBase.Dispose()
            m_dbTransaction.Dispose()
        End If
    End Sub

End Class
