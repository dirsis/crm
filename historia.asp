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
      <div class="page-header">
        <h1>Historia</h1>
<%		
				sql=" select c.apellido,c.nombre,c.razon,g.idgestion,g.idclien,c.razon,c.apellido,c.nombre,g.asunto,g.tema,g.fecha,g.fcompro,g.aquien,g.estado,p.apellido as registro from gestion g, clien c, pass p where c.idclien=g.idclien and g.idpass=p.idpass "
				if request.Form("inputbusca")<>"" then
					sql=sql & " and ( c.apellido like '%" & request.Form("inputbusca") & "%' or c.nombre like '%" & request.Form("inputbusca") & "%' or c.razon like '%" & request.Form("inputbusca") & "%' or g.asunto like '%" & request.Form("inputbusca") & "%' or g.tema like '%" & request.Form("inputbusca") & "%')"
				end if
				sql=sql & " order by g.fecha"
				set rs = oConn.Execute(sql)
%>
      </div>
      <div class="row">
        <div class="col-md-12">
	        <div class="col-md-3">
		<h3>Cliente: <%=request.QueryString("idclien")%></h3>
		<%IF not rs.eof  THEN%>
		<h4>Apellido: <%=rs("apellido")%></h4>
		<h4>Nombre: <%=rs("nombre")%></h4>
		<h4>Razon Social: <%=rs("razon")%></h4>
		<%end if%>
			</div>
	        <div class="col-md-3">
        		<form class="form-signin" action="historia.asp?idclien=<%=request.QueryString("idclien")%>" method="post">
        		<label >Buscador</label><input type="text" id="inputbusca" name="inputbusca" class="form-control" placeholder="Texto por Aprox." required="" autofocus="">
				<button class="btn btn-lg btn-primary btn-block" type="submit">Buscar</button>
				</form>
			</div>
		</div>
		<div class="col-md-12">
          <table class="table" width="100%">
            <thead>
              <tr>
			  <th colspan="13">

			  </th></tr>
			  <tr>
                <th>Asunto</th>
                <th>Tema</th>
                <th>Fecha</th>
                <th>Compromiso</th>
                <th>Estado</th>
                <th>Registro</th>
              </tr>
            </thead>
            <tbody>
			  <%
				sql=" select g.idgestion,g.asunto,g.tema,g.fecha,g.fcompro,g.aquien,g.estado,p.apellido as registro from gestion g, pass p where g.idclien='" & request.QueryString("idclien") & "' and g.idpass=p.idpass "
				if request.Form("inputbusca")<>"" then
					sql=sql & " and ( g.asunto like '%" & request.Form("inputbusca") & "%' or g.tema like '%" & request.Form("inputbusca") & "%')"
				end if
				sql=sql & " order by g.fecha"
				set rs = oConn.Execute(sql)
				IF not rs.eof  THEN
				rs.movefirst
				do while not rs.eof 
					%>
	                <tr>
                	<td><%=rs("asunto")%></td>
                	<td><%=rs("tema")%></td>
                	<td><%=rs("fecha")%></td>
                	<td><%=rs("fcompro")%></td>
                	<td><%=rs("estado")%></td>
                	<td><%=rs("registro")%></td>
              	    </tr>
					<%
					rs.movenext
					loop
				end if
				%>
            </tbody>
          </table>
        </div></div>
    <!-- IE10 viewport hack for Surface/desktop Windows 8 bug -->
    <script src="js/ie10-viewport-bug-workaround.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.3/jquery.min.js"></script>
    <script>window.jQuery || document.write('<script src="js/vendor/jquery.min.js"><\/script>')</script>
    <script src="js/bootstrap.min.js"></script>
    <script src="js/ie10-viewport-bug-workaround.js"></script>
    <script src="js/offcanvas.js"></script>
</body></html>