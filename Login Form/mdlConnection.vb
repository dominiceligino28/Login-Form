Imports System.Data.OleDb
Imports System.Data
Module mdlConnection

    Public connStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\Dominic Mina\Downloads\Database1.mdb"
    Public conn As OleDbConnection = New OleDbConnection(connStr)
        Public sql As String
        Public dt As New DataTable

    Function ExecuteQuery(ByVal Query As String) As DataTable

        Dim sqlDT As New DataTable
        Dim sqlCon As New OleDbConnection(connStr)
        Dim sqlDA As New OleDbDataAdapter(Query, sqlCon)
        Dim sqlCB As New OleDbCommandBuilder(sqlDA)
        sqlDA.Fill(sqlDT)
        Return sqlDT
    End Function
End Module
