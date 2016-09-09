<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Buffer=True%> 
<!-- #include file="odbc.asp" -->
<%
if session("idpass")="" then
	response.Redirect("login.asp")
end if
%>
<!DOCTYPE html>
<html lang="en">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
<meta http-equiv="X-UA-Compatible" content="IE=edge">
<meta name="viewport" content="width=device-width, initial-scale=1">
<meta name="description" content="">
<meta name="author" content="">
<title>Print</title>
<style type="text/css">
<!--
.Estilo1 {font-size: 16px}
-->
</style>
</head>
<body>
<%if request.QueryString("idclien")<>"" then
	idclien=request.QueryString("idclien")
	sql=" select g.idcomprob,g.motivo,g.total,g.fecha,g.nota,p.apellido as registro,c.apellido,c.nombre,c.razon,c.mail1 from erp_comprob g, erp_clien c, pass p where g.idclien=c.idclien and g.idpass=p.idpass and g.idclien='" & idclien & "' order by g.fecha"
	set rs = oConn.Execute(sql)
	IF not rs.eof  THEN
		xcorreo=rs("mail1")
		XTEMA="<table width=600 align=center style=""background-color:#FFFFFF; font-family:Verdana, Arial, Helvetica, sans-serif; font-size:10px; border:1px;"">"
		XTEMA=XTEMA & "<thead><tr><th colspan=3><div align=left><h3><img src=""http://gacela.com.ar/wp-content/uploads/2016/08/cropped-dirsis1-300x300.png"" alt=""Dirsis Desarrollos"" width=""100""></h3></div></th>"
		XTEMA=XTEMA & "<th colspan=2><div align=right><h3>Fecha " & date & "</h3></div></th>"
		XTEMA=XTEMA & "</tr><tr><th colspan=5><div align=left>Apellido:<br>" & rs("apellido") & "," & rs("nombre") & "(" & rs("razon") & ")</div></th>"
		XTEMA=XTEMA & "</tr></thead></table><hr><table width=600 align=center bgcolor=#FFFF99><thead>"
		XTEMA=XTEMA & "<tr style=""background-color:#333333; color: #FFFFFF;"">"
		XTEMA=XTEMA & "<th width=""11%"">Fecha</th><th width=""50%"">Motivo</th><th width=""10%"">Total</th><th width=""10%"">Acumulado</th><th width=""19%"">Nota</th></tr></thead><tbody>"
		rs.movefirst
		xacum=0
		do while not rs.eof 
			xacum=xacum+rs("total")
			XTEMA=XTEMA & "<tr style=""background-color:#CCFFCC; color:#000033;"">"
			XTEMA=XTEMA & "<td>" & rs("fecha") & "</td>"
			XTEMA=XTEMA & "<td>" & rs("motivo") & "</td>"
			XTEMA=XTEMA & "<td>" & FormatNumber(rs("total")) & "</td>"
			XTEMA=XTEMA & "<td>" & FormatNumber(xacum) & "</td>"
			XTEMA=XTEMA & "<td>" & rs("nota") & "</td>"
			XTEMA=XTEMA & "</tr>"
			rs.movenext
		loop
		XTEMA=XTEMA & "</tbody><thead><tr style=""background-color:#333333; color: #FFFFFF; font-size:20px;"">"
		XTEMA=XTEMA & "<th width=""11%""></th><th width=""50%""></th><th width=""10%"">Saldo</th><th width=""10%"">" & FormatNumber(xacum) & "</th>"
		XTEMA=XTEMA & "<th width=""19%""></th>"
		XTEMA=XTEMA & "</tr>			</thead>		</table>		<hr>		<table width=600 align=center>			<thead>"
		XTEMA=XTEMA & "<tr style=""background-color:#FFFFFF; color:#333333; font-size:14px;"">"
		XTEMA=XTEMA & "<th width=""100%"" colspan=5>"
		XTEMA=XTEMA & "			  Los datos informados pueden haber variado por transacciones no procesadas al momento de esta publicacion, si tiene dudas de actualizacion por favor consultar a la empresa.</th>"
		XTEMA=XTEMA & "  <tr style=""background-color:#FFFFFF; color:#333333; font-size:14px;"">"
		XTEMA=XTEMA & "  <th colspan=5>Emitido por software nube <a href=""http://www.gacela.com.ar"">Gacela </a></th></tr>"
  		XTEMA=XTEMA & "</thead></table>"
		call MailSendergmail("DIRSIS DESARROLLOS - Notificacion de Saldos pendientes",xtema,xcorreo)	
		response.Write("se envivo mail a " & xcorreo & "<br>")
		xcorreo="dirsis@gmail.com"
		call MailSendergmail("DIRSIS DESARROLLOS - Notificacion de Saldos pendientes",xtema,xcorreo)	
		response.Write("se envivo mail a " & xcorreo & "<br>")
	end if
end if
%>
</body></html>