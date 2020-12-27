Imports System
Imports MySql.Data.MySqlClient
Public Class conMySQL
    Private Correo As Email
    Private CON As MySqlConnection
    Private DA As MySqlDataAdapter
    Private DS As DataSet
    Private DR As System.Data.DataTableReader
    Public Sub New(ByVal Server As String, ByVal Database As String, ByVal User As String, ByVal Pwd As String)
        Correo = New Email
        Try
            CON = New MySqlConnection("server=" & Server & ";database=" & Database & ";Uid=" & User & ";Pwd=" & Pwd & ";")
            CON.Open()
        Catch ex As MySqlException
            Correo.Enviar("sistemas@tiendadigital.com.ve", "Excepcion de MySQL en New(....)", ex.Message, False)
        End Try
    End Sub
    Public Function Consulta(ByVal SQLquery As String) As System.Data.DataTable
        Try
            DA = New MySqlDataAdapter(SQLquery, CON)
            DS = New DataSet
            DA.Fill(DS)
            DR = DS.Tables(0).CreateDataReader
        Catch ex As MySqlException
            Correo.Enviar("sistemas@tiendadigital.com.ve", "Excepcion de MySQL en Consulta(.)", ex.Message, False)
            Return New System.Data.DataTable
        End Try
        Return DS.Tables(0)
    End Function
    Public Function hayRegistros() As Boolean
        Return DS.Tables(0).CreateDataReader.HasRows
    End Function
    Public Function Siguiente() As Boolean
        Return DR.Read
    End Function
    Public Function Campo(ByVal indice As Integer) As String
        Return DR.GetValue(indice).ToString
    End Function
    Public Sub Sentencia(ByVal SQLquery As String)
        Try
            Dim command As New MySqlCommand(SQLquery, CON)
            command.ExecuteNonQuery()
        Catch ex As MySqlException
            Correo.Enviar("sistemas@tiendadigital.com.ve", "Excepcion de MySQL en Sentencia(.)", ex.Message, False)
        End Try
    End Sub
End Class
