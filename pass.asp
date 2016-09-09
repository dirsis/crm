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
  if request.QueryString("idpass")<>0 then
  		if request.QueryString("accion")="x" then
				sql=" delete from pass where idpass=" &request.QueryString("idpass")
				set rs = oConn.Execute(sql)
		end if
  end if
  %>
	<%
	call menu
	%>

      <div class="page-header">
        <h1>Usuarios</h1>
      </div>
      <div class="row">
        <div class="col-md-12">
   		<form class="form-signin" action="clientes.asp" method="post">
       		<label >Buscador</label><input type="text" id="inputbusca" name="inputbusca" class="form-control" placeholder="Texto por Aprox." required="" autofocus="">
			<button class="btn btn-lg btn-primary btn-block" type="submit">Buscar</button>
		</form>
          <table class="table" width="100%">
            <thead>
              <tr>
			  <th colspan="13">

			  </th></tr>
			  <tr>
                <th>Accion</th>
                <th>Id</th>
                <th>Usuario</th>
                <th>Clave</th>
                <th>Apellido</th>
                <th>Mail</th>
                <th>Nivel</th>
                <th>Localidad</th>
                <th>Provincia</th>
                <th>Telefono</th>
              </tr>
            </thead>
            <tbody>
			  <%
			  	sql=" select * from pass "
				if request.Form("inputbusca")<>"" then
					sql=sql & " where apellido like '%" & request.Form("inputbusca") & "%' or correo like '%" & request.Form("inputbusca") & "%' or mail like '%" & request.Form("inputbusca") & "%'"
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
					    <a href="abcpass.asp?accion=e&idpass=<%=rs("idpass")%>"><span class="label label-warning">E</span></a>
				        <a href="pass.asp?accion=x&idpass=<%=rs("idpass")%>"><span class="label label-danger">X</span></a>
				      </p>					
					</td>
    	            <td><%=rs("idpass")%></td>
        	        <td><%=rs("mail")%></td>
            	    <td><%=rs("clave")%></td>
                	<td><%=rs("apellido")%></td>
                	<td><%=rs("correo")%></td>
                	<td><%=rs("nivel")%></td>
                	<td><%=rs("domicilio")%></td>
                	<td><%=rs("localidad")%></td>
                	<td><%=rs("telefono")%></td>
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