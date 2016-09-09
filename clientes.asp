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

    <!-- Bootstrap core CSS -->
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
  if request.QueryString("idclien")<>0 then
  		if request.QueryString("accion")="x" then
				sql=" delete from clien where idclien=" &request.QueryString("idclien")
				set rs = oConn.Execute(sql)
		end if
  end if
  %>
	<%
	call menu
	%>

      <div class="page-header">
        <h1>Clientes</h1>
      </div>
      <div class="row">
        <div class="col-md-12">
		        		<form class="form-signin" action="clientes.asp" method="post">
        		<label >Buscador</label><input type="text" id="inputbusca" name="inputbusca" class="form-control" placeholder="Texto por Aprox." required="" autofocus="">
				<button class="btn btn-lg btn-primary btn-block" type="submit">Buscar</button>
				</form>
			      <p>
					<span class="label label-success">G) agrega Gestión</span>
					<span class="label label-info">H) Historia</span>
					<span class="label label-warning">E) Editar</span>
					<span class="label label-danger">X) Eliminar</span>
			      </p>	
          <table class="table" width="100%">
            <thead>
              <tr>
			  <th colspan="13">

			  </th></tr>
			  <tr>
                <th>Accion</th>
                <th>#</th>
                <th>Razon Social</th>
                <th>Apellido</th>
                <th>Nombre</th>
                <th>Domicilio</th>
                <th>Localidad</th>
                <th>Provincia</th>
                <th>Telefono1</th>
                <th>Telefono2</th>
                <th>Telefono3</th>
                <th>Mail1</th>
                <th>Mail2</th>
                <th>Mail3</th>
              </tr>
            </thead>
            <tbody>
			  <%
			  	sql=" select * from clien "
				if request.Form("inputbusca")<>"" then
					sql=sql & " where apellido like '%" & request.Form("inputbusca") & "%' or nombre like '%" & request.Form("inputbusca") & "%' or razon like '%" & request.Form("inputbusca") & "%'"
				end if
				sql=sql & " order by apellido "
				set rs = oConn.Execute(sql)
				IF not rs.eof  THEN
				rs.movefirst
				do while not rs.eof 
					%>
	                <tr>
    	            <td>
				      <p>
				        <a href="abcgestion.asp?accion=g&idclien=<%=rs("idclien")%>"><span class="label label-success">G</span></a>
				        <a href="historia.asp?idclien=<%=rs("idclien")%>"><span class="label label-info">H</span></a>
					    <a href="abccli.asp?accion=e&idclien=<%=rs("idclien")%>"><span class="label label-warning">E</span></a>
				        <a href="clientes.asp?accion=x&idclien=<%=rs("idclien")%>"><span class="label label-danger">X</span></a>
				      </p>					
					</td>
    	            <td><%=rs("idclien")%></td>
        	        <td><%=rs("razon")%></td>
            	    <td><%=rs("apellido")%></td>
                	<td><%=rs("nombre")%></td>
                	<td><%=rs("domicilio")%></td>
                	<td><%=rs("localidad")%></td>
                	<td><%=rs("provincia")%></td>
                	<td><%=rs("telefono1")%></td>
                	<td><%=rs("telefono2")%></td>
                	<td><%=rs("telefono3")%></td>
                	<td><%=rs("mail1")%></td>
                	<td><%=rs("mail2")%></td>
                	<td><%=rs("mail3")%></td>
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