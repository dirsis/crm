<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Buffer=True%> 
<!-- #include file="odbc.asp" -->
<%
if session("idpass")="" then
	response.Redirect("login.asp")
end if
%>
<!DOCTYPE html>
<!-- saved from url=(0040)http://getbootstrap.com/examples/signin/ -->
<html lang="en"><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <!-- The above 3 meta tags *must* come first in the head; any other head content must come *after* these tags -->
    <meta name="description" content="">
    <meta name="author" content="">
    <link rel="icon" href="http://getbootstrap.com/favicon.ico">

    <title>CMR</title>

    <link href="css/bootstrap.min.css" rel="stylesheet">
    <!-- IE10 viewport hack for Surface/desktop Windows 8 bug -->
    <link href="css/ie10-viewport-bug-workaround.css" rel="stylesheet">

    <!-- Custom styles for this template -->
    <link href="css/signin.css" rel="stylesheet">

    <!-- Just for debugging purposes. Don't actually copy these 2 lines! -->
    <!--[if lt IE 9]><script src="js/ie8-responsive-file-warning.js"></script><![endif]-->
    <script src="js/ie-emulation-modes-warning.js"></script>

    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->
  </head>

  <body>
	<%
	call menu
	%>
		<form class="form-signin" action="erp_historia.asp?idclien=<%=request.QueryString("idclien")%>" method="post">
		<div class="page-header">
		<h1>Historia</h1>
			<div class="row">
			<div class="col-md-6"></div>
			<div class="col-md-6"><label >Buscador</label><input type="text" id="inputbusca" name="inputbusca" class="form-control" placeholder="Texto por Aprox."  autofocus="">
			<button class="btn btn-lg btn-primary btn-block" type="submit">Buscar</button></div>
			</div>
		</div>
		</form>
      <div class="row">
        <div class="col-md-12">
	        <div class="col-md-3">
			<%if request.QueryString("idclien")<>"" then%>
				<h3>Cliente: <%=request.QueryString("idclien")%></h3> 
				<a href="erp_print.asp?id=historia&idclien=<%=request.QueryString("idclien")%>" target="_blank">version imprimible</a>
				<a href="erp_mail.asp?id=historia&idclien=<%=request.QueryString("idclien")%>" target="_blank">version mail</a>
			  <%if request.QueryString("idclien")<>"" then
					sql="select * from erp_clien where idclien=" & request.QueryString("idclien")
					set rs = oConn.Execute(sql)
					if not rs.eof then%>
					<h4>Apellido: <%=rs("apellido")%></h4>
					<h4>Nombre: <%=rs("nombre")%></h4>
					<h4>Razon Social: <%=rs("razon")%></h4>
					<%end if%>
				<%end if%>
			<%end if%>
			</div>
		</div>
		<div class="col-md-12">
          <table class="table" width="100%">
            <thead>
			  <tr>
                <th>Acciones</th>
                <th>Fecha</th>
                <th>motivo</th>
                <th>Total</th>
                <th>Acumulado</th>
                <th>Cliente</th>
                <th>Nota</th>
                <th>Registro</th>
              </tr>
            </thead>
            <tbody>
			  <%
			  	if request.QueryString("idcomprob")<>0 then
					if request.QueryString("accion")="x" then
						sql="delete from erp_comprob where idcomprob=" & request.QueryString("idcomprob")
						set rs = oConn.Execute(sql)
					end if
					if request.QueryString("accion")="c" then
						sql="insert into erp_comprob (fecha,numero,motivo,total,nota,idclien,idpass) select fecha,numero,motivo,total,nota,idclien," & session("idpass") & " as idpass from erp_comprob where idcomprob=" & request.QueryString("idcomprob")
						set rs = oConn.Execute(sql)
					end if
				end if
				sql=" select g.idcomprob,g.motivo,g.total,g.fecha,g.nota,p.apellido as registro,c.apellido,c.nombre from erp_comprob g, erp_clien c, pass p where g.idclien=c.idclien and g.idpass=p.idpass "
				if request.QueryString("idclien")<>"" then
					sql=sql & " and g.idclien='" & request.QueryString("idclien") & "' "
				end if
				if request.Form("inputbusca")<>"" then
					sql=sql & " and ( g.motivo like '%" & request.Form("inputbusca") & "%' or c.apellido like '%" & request.Form("inputbusca") & "%')"
				end if
				sql=sql & " order by g.fecha"
				set rs = oConn.Execute(sql)
				IF not rs.eof  THEN
				rs.movefirst
				xacum=0
				do while not rs.eof 
					xacum=xacum+rs("total")
					%>
	                <tr>
					<td>
				      <p>
				        <a href="erp_historia.asp?accion=c&idcomprob=<%=rs("idcomprob")%>"><span class="label label-success">C</span></a>
				        <a href="erp_historia.asp?accion=x&idcomprob=<%=rs("idcomprob")%>"><span class="label label-danger">X</span></a>
				        <a href="erp_abcfac.asp?accion=e&idcomprob=<%=rs("idcomprob")%>"><span class="label label-info">E</span></a>
				      </p>					
					</td>					
                	<td><%=rs("fecha")%></td>
                	<td><%=rs("motivo")%></td>
                	<td><%=rs("total")%></td>
                	<td><%=xacum%></td>
                	<td><%=rs("apellido")& "," & rs("nombre")%></td>
                	<td><%=rs("nota")%></td>
                	<td><%=rs("registro")%></td>
              	    </tr>
					<%
					rs.movenext
					loop
				end if
				%>
            </tbody>
          </table>
        </div>
	</div>
    <!-- IE10 viewport hack for Surface/desktop Windows 8 bug -->
    <script src="js/ie10-viewport-bug-workaround.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.3/jquery.min.js"></script>
    <script>window.jQuery || document.write('<script src="js/vendor/jquery.min.js"><\/script>')</script>
    <script src="js/bootstrap.min.js"></script>
    <script src="js/ie10-viewport-bug-workaround.js"></script>
    <script src="js/offcanvas.js"></script>
</body></html>