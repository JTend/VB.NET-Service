Imports System.Data.SqlClient

Public Class conSQLServer
    Private CON As SqlConnection
    Private DA As SqlDataAdapter
    Private DS As DataSet
    Private DR As System.Data.DataTableReader
    Private CoRrEo As Email
    Public Sub New(ByVal Server As String, ByVal Database As String, ByVal User As String, ByVal Pwd As String)
        CoRrEo = New Email
            CON = New SqlConnection("server=" & Server & ";database=" & Database & ";Uid=" & User & ";Pwd=" & Pwd & ";")
            CON.Open()
    End Sub
    Public Function Consulta(ByVal SQLquery As String) As System.Data.DataTable
        DA = New SqlDataAdapter(SQLquery, CON)
        DS = New DataSet
        DA.Fill(DS)
        DR = DS.Tables(0).CreateDataReader
        Return DS.Tables(0)
    End Function
    Public Function hayRegistros() As Boolean
        Return DS.Tables(0).CreateDataReader.HasRows
    End Function
    Public Function Siguiente() As Boolean
        Return DR.Read
    End Function
    Public Function Campo(ByVal id_campo As Integer) As String
        Return DR.GetValue(id_campo).ToString
    End Function
    Public Sub Sentencia(ByVal SQLquery As String)
        DA = New SqlDataAdapter(SQLquery, CON)
    End Sub
End Class
