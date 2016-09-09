<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Buffer=True%> 
<!-- #include file="odbc.asp" -->
<!DOCTYPE html>
<html lang="en"><head><meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <meta name="description" content="">
    <meta name="author" content="">
    <link rel="icon" href="http://getbootstrap.com/favicon.ico">
    <title>CMR</title>
    <link rel="stylesheet" href="css/bootstrap.css">
    <link href="css/bootstrap.min.css" rel="stylesheet">
  </head>
  <body>
	<%
  	if request.form("inputcorreo")<>"" then
	  	sql=" insert into mailing  (correo) values ('" & request.form("inputcorreo") & "')"
		set rs = oConn.Execute(sql)
		msg=sql
	end if
	call menu
	%>
	<div class="container">
	<div class="page-header"><h1>Registro de Informacion de General</h1></div>
    <div class="row row-offcanvas row-offcanvas-right">
      	<form  action="abcexpress.asp" method="post">
        <div class="col-xs-12 col-sm-9" id="primero" name="primero">
        <div class="col-xs-8"><label id="linputcorreo" >Correo Electronico (*)</label>
		<input type="mail" id="inputcorreo" name="inputcorreo" value="<%=xcorreo%>" class="form-control" required="" autofocus="">
		</div>
		<div class="col-xs-12 col-sm-9" id="primero" name="primero">
		<button class="btn btn-lg btn-primary btn-block" type="submit" id="btnterminar" name="btnterminar" value="ok">Terminar</button>
		</div>
		<div class="col-xs-12 col-sm-9" id="primero" name="primero">
				<%=msg%>
		</div>
		</div>		
      	</form>
	</div>
	</div>
    <script src="js/bootstrap.min.js"></script>
	<script src="js/jquery.min.js"></script>
</body>
</html>
