Imports System.Data
Imports System.Data.SqlClient

Public Class FsDataBase

    Private objCmd As SqlCommand
    Private objConn As SqlConnection
    Private adData As SqlDataAdapter
    Private Trans As SqlTransaction

    Private iRow As Integer = 0
    Private _Datatable As DataTable

    Private UseTransaction As Boolean = False
    Public HasError As Boolean = False

    Public ErrorMessage As String
    Dim strConnString As String

    Enum DataType
        adVarChar = 0
        adInteger = 1
        adFloat = 2
        adDate = 3
        adDateTime = 3
        adNVarChar = 4
    End Enum

    Enum Direction
        adInput = 0
        adOutput = 1
    End Enum

    Public Sub Close()
        objConn.Close()
        objConn.Dispose()
    End Sub

    Public Sub Reopen()
        Try
            objConn.Close()
            objConn.Open()

        Catch ex As Exception
            HasError = True
            ErrorMessage = ex.Message

        End Try
    End Sub

    Public Sub BeginTransaction()
        Trans = objConn.BeginTransaction(IsolationLevel.ReadCommitted)
        UseTransaction = True
    End Sub

    Public Sub CommitTransaction()
        Trans.Commit()
        UseTransaction = False
    End Sub

    Public Sub Rollback()
        Trans.Rollback()
        UseTransaction = False
    End Sub

    Public Sub New()
        strConnString = ConfigurationManager.AppSettings("FsConnection")

        Try
            HasError = False
            objConn = New SqlConnection(strConnString)
            objConn.Open()

        Catch ex As Exception
            HasError = True
            ErrorMessage = ex.Message

        End Try
    End Sub

    Public Function Execute(ByVal strSQL As String) As Boolean
        objCmd = New SqlCommand

        If UseTransaction Then
            objCmd.Transaction = Trans
        End If

        objCmd.CommandText = strSQL
        objCmd.Connection = objConn

        HasError = False

        Try
            objCmd.ExecuteNonQuery()

        Catch ex As Exception
            HasError = True
            ErrorMessage = ex.Message

        End Try

        objCmd = Nothing
    End Function

    Public Sub OpenTable(ByVal strsql As String)
        iRow = 0
        _Datatable = New DataTable

        Try
            HasError = False

            adData = New SqlDataAdapter(strsql, objConn)
            adData.SelectCommand.CommandTimeout = 300

            If UseTransaction Then
                adData.SelectCommand.Transaction = Trans
            End If

            adData.Fill(_Datatable)
            adData = Nothing

        Catch ex As Exception
            HasError = True
            ErrorMessage = ex.Message

        End Try
    End Sub

    Public Sub MoveNext()
        iRow = iRow + 1
    End Sub

    Public Sub MovePrevious()
        If iRow <> 0 Then
            iRow = iRow - 1
        End If
    End Sub

    Public Sub Last()
        iRow = _Datatable.Rows.Count - 1
    End Sub

    Public Sub First()
        iRow = 0
    End Sub

    Public Function Item(ByVal sItem As String) As String
        Try
            Return Trim(_Datatable.Rows(iRow).Item(sItem))

        Catch ex As Exception
            Return ""

        End Try
    End Function

    Public Function Bof() As Boolean
        If iRow = 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function Eof() As Boolean
        Try
            If iRow = _Datatable.Rows.Count Then
                Return True
            Else
                Return False
            End If

        Catch ex As Exception
            Return True

        End Try
    End Function

    Public Function getDataSet(ByVal strSQL As String) As DataSet
        Dim ds As DataSet
        ds = New DataSet

        Try
            HasError = False
            adData = New SqlDataAdapter(strSQL, objConn)
            adData.Fill(ds)
            adData = Nothing

        Catch ex As Exception
            HasError = True
            ErrorMessage = ex.Message

        End Try

        Return ds
    End Function

    Public Function getDataReader(ByVal strsql As String) As SqlDataReader
        Dim dtReader As SqlDataReader

        objCmd = New SqlCommand(strsql, objConn)
        dtReader = objCmd.ExecuteReader()

        Return dtReader
    End Function

    Public Function Datasource() As DataTable
        Return _Datatable
    End Function

    Public Function RecordCount() As Integer
        Return _Datatable.Rows.Count
    End Function

    Protected Overrides Sub Finalize()
        objConn = Nothing
        MyBase.Finalize()
    End Sub
End Class
