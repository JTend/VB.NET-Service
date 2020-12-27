Imports System.Threading

Public Class enviarVentas
    Public Shared Sistema As conMySQL
    Private Tempo As System.Threading.Timer
    Private Tiendas As Interfaz()
    Private Correo As Email
    Public Sub New()
        InitializeComponent()
        Eventos.Source = "Ventas"
        Eventos.Log = ""
    End Sub
    Protected Overrides Sub OnStart(ByVal args() As String)
        Correo = New Email
        Dim Proceso As New TimerCallback(AddressOf alterInicio)
        Tempo = New System.Threading.Timer(Proceso, Nothing, 0, 60000)
        Eventos.WriteEntry("Servicio Iniciado")
    End Sub

    Protected Overrides Sub OnStop()
        Eventos.WriteEntry("Servicio Detenido")
    End Sub
    Private Sub alterInicio()
        If Correo.Estado = 0 Then
            Correo.Estado = 1
            Correo.Enviar("dtosistemas@tiendadigital.com.ve", "Se ha iniciado el servicio emailServidigital", "Correo.Estado = " & Correo.Estado.ToString, False)
        Else
            Inicio()
        End If
    End Sub
    Private Sub Inicio()
        Sistema = New conMySQL("SERVERMYSQL", "codica2012", "root", "password")
        Sistema.Consulta("select descrip from horas where hora <= curtime() and hora > (select ifnull(max(time(hora)),'00:00') from enviados where date(hora) = curdate()) order by hora desc")
        If Sistema.Siguiente Then
            Try
                Correo.Enviar("Resúmen: " & Sistema.Campo(0), mensaje(), True)
                Sistema.Sentencia("insert into enviados (hora) values (current_timestamp())")
            Catch exe As Exception
                Correo.Enviar("dtosistemas@tiendadigital.com.ve", "Excepcion sistemVentas" & exe.ToString, exe.Message, False)
            End Try
        End If
    End Sub

    Private Function mensaje() As String
        Dim Tabla As New System.Data.DataTable
        Sistema = New conMySQL("SERVERMYSQL", "servicios", "root", "password")
        Tabla = Sistema.Consulta("select * from tienda")
        Dim i As Integer = 0
        ReDim Tiendas(Tabla.Rows.Count - 1)
        While Sistema.Siguiente
            Tiendas(i) = New Interfaz(Sistema.Campo(1), Sistema.Campo(2), Sistema.Campo(3), Sistema.Campo(4), Sistema.Campo(5), Sistema.Campo(6), Integer.Parse(Sistema.Campo(7)))
            i = i + 1
        End While
        Dim T As Interfaz
        Dim msg As String
        Dim tW, tX, tY, tZ As Double
        tW = 0
        tX = 0
        tY = 0
        tZ = 0
        msg = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">" _
        & "<html>" _
        & "<head>" _
        & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">" _
        & "<title>Resumen Comercial</title>" _
        & "</head>" _
        & "<body>" _
        & "<table border = 1 cellspacing=0 cellpadding=2 bordercolor=""blue"">" _
        & "<caption>Relación Tiendas Codica</caption>" _
        & "<tr>" _
        & "<th colspan=""7""; bgcolor=cyan>Ventas de Contado</th>" _
        & "</tr>" _
        & "<tr>" _
        & "<th>Tienda</th><th>Efectivo Dia</th><th>Total Dia</th><th>Total Mes</th><th>Meta</th><th>Por lograr</th><th>Logrado</th>" _
        & "</tr>"
        'INSERCIÓN DE VENTAS DE CONTADO
        For Each T In Tiendas
            Dim W, X, Y, Z As Double
            W = T.efectivoDia
            X = T.ventasContadoDia
            Y = T.ventasContadoMes
            Z = T.metaContado
            tW = tW + W
            tX = tX + X
            tY = tY + Y
            tZ = tZ + Z
            If T.conectado Then
                msg = msg & "<tr><td>"
            Else
                msg = msg & "<tr><td><FONT COLOR=""red"">"
            End If
            msg = msg & T.getNombre & "</td><td align=""right"">" _
                & FormatNumber(W.ToString, 2) & "</td><td align=""right"">" _
                & FormatNumber(X.ToString, 2) & "</td><td align=""right"">" _
                & FormatNumber(Y.ToString, 2) & "</td><td align=""right"">" _
                & FormatNumber(Z.ToString, 2) & "</td><td align=""right"">"
            If (Z - Y) < 0 Then
                msg = msg & "0" & "</td><td align=""right"">"
            Else
                msg = msg & FormatNumber((Z - Y).ToString, 2) & "</td><td align=""right"">"
            End If
            msg = msg & FormatNumber((Y * 100 / Z).ToString, 2) & "%" & "</td></tr>"
        Next
        msg = msg & "<tr bgcolor=""yellow""><th>Totales</th><th align=""right"">" & FormatNumber(tW.ToString, 2) & "</th><th align=""right"">" & FormatNumber(tX.ToString, 2) & "</th><th align=""right"">" & FormatNumber(tY.ToString, 2) & "</th><th align=""right"">" & FormatNumber(tZ.ToString) & "</th><th align=""right"">"
        If (tZ - tY) < 0 Then
            msg = msg & "0" & "</th><th align=""right"">"
        Else
            msg = msg & FormatNumber((tZ - tY).ToString, 2) & "</th><th align=""right"">"
        End If
        msg = msg & FormatNumber((tY * 100 / tZ).ToString, 2) & "%" & "</th></tr>" _
        & "<tr><th colspan=""6""; bgcolor=cyan>Ventas a Crédito (Credihogar)</th></tr>" _
        & "<tr>" _
        & "<th>Tienda</th><th>Hoy</th><th>Este Mes</th><th>Meta</th><th>Por lograr</th><th>Logrado</th>" _
        & "</tr>"
        tX = 0
        tY = 0
        tZ = 0
        'INSERCIÓN DE VENTAS A CREDITO
        For Each T In Tiendas
            Dim X, Y, Z As Double
            X = T.ventasCreditoDia
            Y = T.ventasCreditoMes
            Z = T.metaCredito
            tX = tX + X
            tY = tY + Y
            tZ = tZ + Z
            If T.conectado Then
                msg = msg & "<tr><td>"
            Else
                msg = msg & "<tr><td><FONT COLOR=""red"">"
            End If
            msg = msg & T.getNombre & "</td><td align=""right"">" _
                & FormatNumber(X.ToString, 2) & "</td><td align=""right"">" _
                & FormatNumber(Y.ToString, 2) & "</td><td align=""right"">" _
                & FormatNumber(Z.ToString, 2) & "</td><td align=""right"">"
            If (Z - Y) < 0 Then
                msg = msg & "0" & "</td><td align=""right"">"
            Else
                msg = msg & FormatNumber((Z - Y).ToString, 2) & "</td><td align=""right"">"
            End If
            msg = msg & FormatNumber((Y * 100 / Z).ToString, 2) & "%" & "</td></tr>"
        Next
        msg = msg & "<tr bgcolor=""yellow""><th>Total Creditos</th><th align=""right"">" & FormatNumber(tX.ToString, 2) & "</th><th align=""right"">" & FormatNumber(tY.ToString, 2) & "</th><th align=""right"">" & FormatNumber(tZ.ToString, 2) & "</th><th align=""right"">"
        If (tZ - tY) < 0 Then
            msg = msg & "0" & "</th><th align=""right"">"
        Else
            msg = msg & FormatNumber((tZ - tY).ToString, 2) & "</th><th align=""right"">"
        End If
        msg = msg & FormatNumber((tY * 100 / tZ).ToString, 2) & "%" & "</th></tr></table>"

        tX = 0
        tY = 0
        'VENTAS POR ESTACIONES DE TRABAJO REGISTRADAS
        For Each T In Tiendas
            Dim X As Double
            Dim Y As Double
            Sistema.Consulta("select descrip, codesta from estaciones, tienda where id_tienda = tienda and nombre = '" & T.getNombre & "'")
            If Sistema.Siguiente Then
                Dim sta As String = ""
                Dim prim As Boolean = True
                msg = msg & "<table border = 1 cellspacing=0 cellpadding=2 bordercolor=""blue"">" _
                    & "<caption>Desglose de " & T.getNombre & "</caption>" _
                    & "<tr>" _
                    & "<th colspan=""3""; bgcolor=cyan>Ventas de Contado</th>" _
                    & "</tr>" _
                    & "<tr>" _
                    & "<th>Estacion</th><th>Hoy</th><th>Este Mes</th>" _
                    & "</tr>"
                Do
                    If prim Then
                        sta = "'" & Sistema.Campo(1) & "'"
                        prim = False
                    Else
                        sta = sta & ",'" & Sistema.Campo(1) & "'"
                    End If
                    X = T.ventaEstacion(Sistema.Campo(1))
                    Y = T.ventaAcumEstacion(Sistema.Campo(1))
                    tX = tX + X
                    tY = tY + Y
                    msg = msg & "<tr><th>" & Sistema.Campo(0) & "</th><td align=""right"">" & FormatNumber(X.ToString, 2) & "</td><td align=""right"">" & FormatNumber(Y.ToString, 2) & "</td></tr>"
                Loop While Sistema.Siguiente
                X = T.ventasRestantes(sta)
                Y = T.ventaAcumRestantes(sta)
                tX = tX + X
                tY = tY + Y
                msg = msg & "<tr><th>Otras Estaciones</th><td align=""right"">" & FormatNumber(X.ToString, 2) & "</td><td align=""right"">" & FormatNumber(Y.ToString, 2) & "</td></tr>" _
                    & "<tr bgcolor=yellow><th>Total</th><td align=""right"">" & FormatNumber(tX.ToString, 2) & "</td><td align=""right"">" & FormatNumber(tY.ToString, 2) & "</td></tr></table>" _
                    & "<p><b>Departamento de Sistemas</b></p>"
            End If
        Next
        msg = msg & "</body></html>"
        Return msg
    End Function

    Private Sub Eventos_EntryWritten(ByVal sender As System.Object, ByVal e As System.Diagnostics.EntryWrittenEventArgs) Handles Eventos.EntryWritten

    End Sub
End Class
