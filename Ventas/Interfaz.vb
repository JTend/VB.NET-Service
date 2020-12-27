Imports System.Data.SqlClient

Public Class Interfaz
    Private conTienda As conSQLServer
    Private nombre As String
    Private servidor As String
    Private usuario As String
    Private clave As String
    Private based As String
    Private deposito As String
    Public conectado As Boolean
    Private ImVA As Integer
    Private Sistema As conMySQL
    Private CoRrEo As Email
    Public Sub New(ByVal nom As String, ByVal ser As String, ByVal usu As String, ByVal pwd As String, ByVal bda As String, ByVal dep As String, ByVal iva As Integer)
        Sistema = New conMySQL("SERVERMYSQL", "codica2012", "root", "password")
        CoRrEo = New Email
        nombre = nom
        servidor = ser
        usuario = usu
        clave = pwd
        based = bda
        deposito = dep
        ImVA = iva
        Try
            conTienda = New conSQLServer(servidor, based, usuario, clave)
            conectado = True
        Catch ex As SqlClient.SqlException
            servidor = "SERVERMSSQL"
            based = "codicappal"
            usuario = "sqluser"
            clave = "password"
            Try
                conTienda = New conSQLServer(servidor, based, usuario, clave)
            Catch e As SqlException
                CoRrEo.Enviar("sistemas@tiendadigital.com.ve", "Err " & deposito & "(" & servidor & "): Conexion", e.Message, False)
            End Try
            conectado = False
        End Try
    End Sub
    Public Function getNombre() As String
        Return nombre
    End Function
    Public Function clausulaIVA() As String
        Select Case ImVA
            Case 1
                Return "1.12"
            Case 2
                Return "(case P.esexento when 0 then 1.12 else 1 end)"
            Case Else
                Return "1"
        End Select
    End Function
    Public Sub leerContado()
        Dim Query As String = ""
        Query = ""
    End Sub
    Public Function efectivoDia() As Double
        Try
            conTienda.Consulta("SELECT     ISNULL(SUM(F.CancelE * (CASE F.tipofac WHEN 'A' THEN 1 WHEN 'B' THEN - 1 ELSE 0 END)),0) AS Venta " _
                & "FROM         SAFACT as F " _
                & "WHERE     (F.FechaE >= CONVERT(date, GETDATE())) AND (F.Contado > 0) and F.codubic = '" & deposito & "'")
            If conTienda.Siguiente Then
                Return Double.Parse(conTienda.Campo(0))
            Else
                Return 0
            End If
        Catch ex As SqlException
            CoRrEo.Enviar("sistemas@tiendadigital.com.ve", "Err " & deposito & "(" & servidor & "): efectivoDia()", ex.Message, False)
            Return 0
        End Try
    End Function
    Public Function ventasContadoDia() As Double
        Try
            conTienda.Consulta("SELECT     ISNULL(SUM(F.MtoTotal * (CASE F.tipofac WHEN 'A' THEN 1 WHEN 'B' THEN - 1 ELSE 0 END)),0) AS Venta " _
                & "FROM         SAFACT as F " _
                & "WHERE     (F.FechaE >= CONVERT(date, GETDATE())) AND (F.Contado > 0) and F.codubic = '" & deposito & "'")
            If conTienda.Siguiente Then
                Return Double.Parse(conTienda.Campo(0))
            Else
                Return 0
            End If
        Catch ex As SqlException
            CoRrEo.Enviar("sistemas@tiendadigital.com.ve", "Err " & deposito & "(" & servidor & "): ventasContadoDia()", ex.Message, False)
            Return 0
        End Try
    End Function
    Public Function ventasCreditoDia() As Double
        Try
            conTienda.Consulta("SELECT     ISNULL(SUM(F.MtoTotal * (CASE F.tipofac WHEN 'A' THEN 1 WHEN 'B' THEN - 1 ELSE 0 END)),0) AS Venta" _
                & " FROM         SAFACT as F " _
                & "WHERE     (F.FechaE >= CONVERT(date, GETDATE())) AND (F.Credito > 0) and F.codubic = '" & deposito & "'")
            If conTienda.Siguiente Then
                Return Double.Parse(conTienda.Campo(0))
            Else
                Return 0
            End If
        Catch ex As SqlException
            CoRrEo.Enviar("sistemas@tiendadigital.com.ve", "Err " & deposito & "(" & servidor & "): ventasCreditoDia()", ex.Message, False)
            Return 0
        End Try
    End Function
    Public Function ventasContadoMes() As Double
        Try
            conTienda.Consulta("SELECT     ISNULL(SUM(F.MtoTotal * (CASE F.tipofac WHEN 'A' THEN 1 WHEN 'B' THEN - 1 ELSE 0 END)),0) AS Venta" _
                & " FROM         SAFACT as F " _
                & "WHERE     (F.FechaE >= '1-' + STR(MONTH(getdate())) + '-' + str(year(GETDATE()))) AND (F.Contado > 0) and F.codubic = '" & deposito & "'")
            If conTienda.Siguiente Then
                Return Double.Parse(conTienda.Campo(0))
            Else
                Return 0
            End If
        Catch ex As SqlException
            CoRrEo.Enviar("sistemas@tiendadigital.com.ve", "Err " & deposito & "(" & servidor & "): ventasContadoMes()", ex.Message, False)
            Return 0
        End Try
    End Function
    Public Function ventasCreditoMes() As Double
        Try
            conTienda.Consulta("SELECT     ISNULL(SUM(F.MtoTotal * (CASE F.tipofac WHEN 'A' THEN 1 WHEN 'B' THEN - 1 ELSE 0 END)),0) AS Venta" _
                & " FROM         SAFACT as F " _
                & "WHERE     (F.FechaE >= '1-' + STR(MONTH(getdate())) + '-' + str(year(GETDATE()))) AND (F.Credito > 0) and F.codubic = '" & deposito & "'")
            If conTienda.Siguiente Then
                Return Double.Parse(conTienda.Campo(0))
            Else
                Return 0
            End If
        Catch ex As SqlException
            CoRrEo.Enviar("sistemas@tiendadigital.com.ve", "Err " & deposito & "(" & servidor & "): ventasCreditoMes()", ex.Message, False)
            Return 0
        End Try
    End Function
    Public Function metaContado() As Double
        Sistema.Consulta("select monto from meta, tienda where mes = month(curdate()) and anio = year(curdate()) and tienda = id_tienda and escontado = 1 and nombre = '" & nombre & "'")
        If Sistema.Siguiente Then
            Return Double.Parse(Sistema.Campo(0))
        Else
            Sistema.Consulta("select monto from meta, tienda where tienda = id_tienda and escontado = 1 and nombre = '" & nombre & "' order by id_meta desc")
            If Sistema.Siguiente Then
                Return Double.Parse(Sistema.Campo(0))
            Else
                Return 0
            End If
        End If
    End Function
    Public Function metaCredito() As Double
        Sistema.Consulta("select monto from meta, tienda where mes = month(curdate()) and anio = year(curdate()) and tienda = id_tienda and escontado = 0 and nombre = '" & nombre & "'")
        If Sistema.Siguiente Then
            Return Double.Parse(Sistema.Campo(0))
        Else
            Sistema.Consulta("select monto from meta, tienda where tienda = id_tienda and escontado = 0 and nombre = '" & nombre & "' order by id_meta desc")
            If Sistema.Siguiente Then
                Return Double.Parse(Sistema.Campo(0))
            Else
                Return 0
            End If
        End If
    End Function
    Public Function ventaEstacion(ByVal Estacion As String) As Double
        Try
            conTienda.Consulta("SELECT     ISNULL(SUM(F.MtoTotal * (CASE F.tipofac WHEN 'A' THEN 1 WHEN 'B' THEN - 1 ELSE 0 END)),0) AS Venta" _
                & " FROM         SAFACT as F " _
                & "WHERE     (F.FechaE >= CONVERT(date, GETDATE())) AND (F.codesta = '" & Estacion & "') AND (F.Contado > 0) and F.codubic = '" & deposito & "'")
            If conTienda.Siguiente Then
                Return Double.Parse(conTienda.Campo(0))
            End If
        Catch ex As SqlException
            CoRrEo.Enviar("sistemas@tiendadigital.com.ve", "Err " & deposito & "(" & servidor & "): ventaEstacion(" & Estacion & ")", ex.Message, False)
        End Try
        Return 0
    End Function
    Public Function ventasRestantes(ByVal Estaciones As String) As Double
        Try
            conTienda.Consulta("SELECT     ISNULL(SUM(F.MtoTotal * (CASE F.tipofac WHEN 'A' THEN 1 WHEN 'B' THEN - 1 ELSE 0 END)),0) AS Venta" _
                & " FROM         SAFACT as F " _
                & "WHERE     (F.FechaE >= CONVERT(date, GETDATE())) AND F.codesta not in (" & Estaciones & ") AND (F.Contado > 0) and F.codubic = '" & deposito & "'")
            If conTienda.Siguiente Then
                Return Double.Parse(conTienda.Campo(0))
            End If
        Catch ex As SqlException
            CoRrEo.Enviar("sistemas@tiendadigital.com.ve", "Err " & deposito & "(" & servidor & "): ventasRestantes(" & Estaciones & ")", ex.Message, False)
        End Try
        Return 0
    End Function
    Public Function ventaAcumEstacion(ByVal Estacion As String) As Double
        Try
            conTienda.Consulta("SELECT     ISNULL(SUM(F.MtoTotal * (CASE F.tipofac WHEN 'A' THEN 1 WHEN 'B' THEN - 1 ELSE 0 END)),0) AS Venta" _
                & " FROM         SAFACT as F " _
                & "WHERE     (F.FechaE >= '1-' + STR(MONTH(getdate())) + '-' + str(year(GETDATE()))) AND (F.codesta = '" & Estacion & "') AND (F.Contado > 0) and F.codubic = '" & deposito & "'")
            If conTienda.Siguiente Then
                Return Double.Parse(conTienda.Campo(0))
            End If
        Catch ex As SqlException
            CoRrEo.Enviar("sistemas@tiendadigital.com.ve", "Err " & deposito & "(" & servidor & "): ventaEstacion(" & Estacion & ")", ex.Message, False)
        End Try
        Return 0
    End Function
    Public Function ventaAcumRestantes(ByVal Estaciones As String) As Double
        Try
            conTienda.Consulta("SELECT     ISNULL(SUM(F.MtoTotal * (CASE F.tipofac WHEN 'A' THEN 1 WHEN 'B' THEN - 1 ELSE 0 END)),0) AS Venta" _
                & " FROM         SAFACT as F " _
                & "WHERE     (F.FechaE >= '1-' + STR(MONTH(getdate())) + '-' + str(year(GETDATE()))) AND F.codesta not in (" & Estaciones & ") AND (F.Contado > 0) and F.codubic = '" & deposito & "'")
            If conTienda.Siguiente Then
                Return Double.Parse(conTienda.Campo(0))
            End If
        Catch ex As SqlException
            CoRrEo.Enviar("sistemas@tiendadigital.com.ve", "Err " & deposito & "(" & servidor & "): ventasRestantes(" & Estaciones & ")", ex.Message, False)
        End Try
        Return 0
    End Function
End Class