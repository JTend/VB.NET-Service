Imports System.Net
Imports System.Net.Mail.MailMessage
Public Class Email
    Private Correo As System.Net.Mail.MailMessage
    Private Servidor As System.Net.Mail.SmtpClient
    Public Estado As Integer
    Public Sub New()
        Estado = 0
        Correo = New System.Net.Mail.MailMessage()
        Servidor = New System.Net.Mail.SmtpClient
        Servidor.Host = "smtp.gmail.com"
        Servidor.Port = 587
        Servidor.EnableSsl = True
        Servidor.Credentials = New System.Net.NetworkCredential("someaddress@gmail.com", "password")
    End Sub

    Public Sub Enviar(ByVal txtDest As String, ByVal txtAsunto As String, ByVal txtMensaje As String, ByVal esHTML As Boolean)
        Dim Correo As New System.Net.Mail.MailMessage()
        Correo.From = New System.Net.Mail.MailAddress("someaddress@gmail.com", "Sistema Automatizado")
        Correo.Subject = txtAsunto
        Correo.To.Add(txtDest)
        Correo.IsBodyHtml = esHTML
        Correo.Body = txtMensaje
        Try
            Servidor.Send(Correo)
        Catch ex As System.Net.Mail.SmtpException
            Throw ex
        End Try
    End Sub

    Public Sub Enviar(ByVal txtAsunto As String, ByVal txtMensaje As String, ByVal esHTML As Boolean)
        Dim Correo As New System.Net.Mail.MailMessage()
        Correo.From = New System.Net.Mail.MailAddress("someaddress@gmail.com", "Sistema Automatizado")
        Correo.Subject = txtAsunto
        Correo.To.Add("ventas@tiendadigital.com.ve")
        Correo.IsBodyHtml = esHTML
        Correo.Body = txtMensaje
        Try
            Servidor.Send(Correo)
        Catch ex As System.Net.Mail.SmtpException
            Throw ex
        End Try
    End Sub
    Public Function generarPorVencer(ByVal Tabla As Windows.Forms.DataGridView, ByVal Cedula As String, ByVal Cuenta As String, ByVal Tipo As String, ByVal Cuotas As String, ByVal Monto As String, ByVal Total As String, ByVal Param As String) As String
        Dim Mensaje As String = ""
        Mensaje = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">" _
        & "<html>" _
        & "<head>" _
        & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">" _
        & "<title>Solicitud de Credito</title>" _
        & "</head>" _
        & "<body>" _
        & "<table border = 1>" _
        & "<caption>Solicitud de Credito " & Tipo & "</caption>" _
        & "<tr>" _
        & "<th colspan=""4""; bgcolor=cyan>Productos solicitados</th>" _
        & "</tr>" _
        & "<tr>" _
        & "<th>Codigo</th><th>Descripcion</th><th>Cantidad</th><th>Precio</th>" _
        & "</tr>"
        For i As Int16 = 0 To Tabla.RowCount - 1 Step 1
            Mensaje = Mensaje & "<tr><td>" _
            & Tabla.Rows(i).Cells(0).Value.ToString & "</td><td>" _
            & Tabla.Rows(i).Cells(1).Value.ToString & "</td><td>" _
            & Tabla.Rows(i).Cells(2).Value.ToString & "</td><td>" _
            & Tabla.Rows(i).Cells(3).Value.ToString & "</td><tr>"
        Next
        Mensaje = Mensaje & "<tr><th colspan=""4""; bgcolor=cyan>Datos Financieros</th></tr>" _
        & "<tr><th colspan=""3"">Cedula Cliente:</th><td>" & Cedula & "</td></tr>" _
        & "<tr><th colspan=""3"">Numero Cuenta:</th><td >" & Cuenta & "</td></tr>" _
        & "<tr><th colspan=""3"">Cantidad de Cuotas:</th><td>" & Cuotas & "</td></tr>" _
        & "<tr><th colspan=""3"">Monto por cuota:</th><td >" & Monto & "</td></tr>" _
        & "<tr><th colspan=""3"">Total a pagar:</th><th>" & Total & "</th></tr>" _
        & "<tr><th colspan=""4""; bgcolor=cyan>Version: 2.0 (" & Param & ")</th></tr>" _
        & "</body></html>"
        Return Mensaje
    End Function
    Public Function generarVencidos(ByVal Tabla As Windows.Forms.DataGridView, ByVal Cedula As String, ByVal Nombre As String, ByVal Apellido As String, ByVal Tipo As String, ByVal Cuotas As String, ByVal Monto As String, ByVal Total As String, ByVal Param As String) As String
        Dim Mensaje As String = ""
        Mensaje = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.01 Transitional//EN"">" _
        & "<html>" _
        & "<head>" _
        & "<meta http-equiv=""Content-Type"" content=""text/html; charset=UTF-8"">" _
        & "<title>Solicitud de Credito</title>" _
        & "</head>" _
        & "<body>" _
        & "<table border = 1>" _
        & "<caption>Solicitud de Credito " & Tipo & "</caption>" _
        & "<tr>" _
        & "<th colspan=""4""; bgcolor=cyan>Productos solicitados</th>" _
        & "</tr>" _
        & "<tr>" _
        & "<th>Codigo</th><th>Descripcion</th><th>Cantidad</th><th>Precio</th>" _
        & "</tr>"
        For i As Int16 = 0 To Tabla.RowCount - 1 Step 1
            Mensaje = Mensaje & "<tr><td>" _
            & Tabla.Rows(i).Cells(0).Value.ToString & "</td><td>" _
            & Tabla.Rows(i).Cells(1).Value.ToString & "</td><td>" _
            & Tabla.Rows(i).Cells(2).Value.ToString & "</td><td>" _
            & Tabla.Rows(i).Cells(3).Value.ToString & "</td><tr>"
        Next
        Mensaje = Mensaje & "<tr><th colspan=""4""; bgcolor=cyan>Datos Financieros</th></tr>" _
        & "<tr><th colspan=""3"">Cedula:</th><td>" & Cedula & "</td></tr>" _
        & "<tr><th colspan=""3"">Cliente:</th><td >" & Nombre & " " & Apellido & "</td></tr>" _
        & "<tr><th colspan=""3"">Cantidad de Cuotas:</th><td>" & Cuotas & "</td></tr>" _
        & "<tr><th colspan=""3"">Monto por cuota:</th><td >" & Monto & "</td></tr>" _
        & "<tr><th colspan=""3"">Total a pagar:</th><th>" & Total & "</th></tr>" _
        & "<tr><th colspan=""4""; bgcolor=cyan>Version: 2.0 (" & Param & ")</th></tr>" _
        & "</body></html>"
        Return Mensaje
    End Function
End Class
