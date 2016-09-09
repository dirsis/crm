<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Buffer=True%> 
<!-- #include file="odbc.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Frameset//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-frameset.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>CMR</title>
<meta name="viewport" content="width=device-width, initial-scale=1.0">
    <!-- Bootstrap core CSS -->
    <link href="css/bootstrap.min.css" rel="stylesheet">
	<link href="css/bootstrap-theme.min.css" rel="stylesheet">

    <!-- IE10 viewport hack for Surface/desktop Windows 8 bug -->
    <link href="css/ie10-viewport-bug-workaround.css" rel="stylesheet">

    <!-- Custom styles for this template -->
    <link href="css/offcanvas.css" rel="stylesheet">
	<link href="css/grid.css" rel="stylesheet">
	<link href="css/theme.css" rel="stylesheet">
    <!-- Just for debugging purposes. Don't actually copy these 2 lines! -->
    <!--[if lt IE 9]><script src="../../assets/js/ie8-responsive-file-warning.js"></script><![endif]-->
    <script src="js/ie-emulation-modes-warning.js"></script>
    <!-- HTML5 shim and Respond.js for IE8 support of HTML5 elements and media queries -->
    <!--[if lt IE 9]>
      <script src="https://oss.maxcdn.com/html5shiv/3.7.2/html5shiv.min.js"></script>
      <script src="https://oss.maxcdn.com/respond/1.4.2/respond.min.js"></script>
    <![endif]-->

</head>
<body>
<%
if session("mail")="" then
	session.abandon
	session("mail")=""
	Response.Cookies("mail")=""
	response.Redirect("login.asp")
end if
if request.Cookies("mail")<>"" then
	sql="select * from pass where mail='" & request.Cookies("mail") & "' "
	set rs = oConn.Execute(sql)     
	if not rs.eof then
		session("mail")=rs("mail")
		Response.Cookies("mail")=rs("mail")
		session("apellido")=rs("apellido")
		session("nivel")=rs("nivel")
	end if
end if
if request.QueryString("tipo")="closesession" then
	session.abandon
	session("mail")=""
	Response.Cookies("mail")=""=""
	response.Redirect("index.asp")
end if
	call menu
	sql="select count(idclien) as cantidad from clien"
	set rs = oConn.Execute(sql)
	xclien=rs("cantidad")
	sql="select count(idgestion) as cantidad from gestion where estado=0"
	set rs = oConn.Execute(sql)
	xgestion0=rs("cantidad")
	sql="select count(idpass) as cantidad from pass"
	set rs = oConn.Execute(sql)
	xusuario0=rs("cantidad")
	sql="select count(idclub) as cantidad from club"
	set rs = oConn.Execute(sql)
	xclub0=rs("cantidad")
	sql="select count(idmailing) as cantidad from mailing"
	set rs = oConn.Execute(sql)
	xmailing0=rs("cantidad")
	sql="select count(idgestion) as cantidad from gestion where estado=1"
	set rs = oConn.Execute(sql)
	xgestion1=rs("cantidad")
	sql="select count(idgestion) as cantidad from gestion where estado=2"
	set rs = oConn.Execute(sql)
	xgestion2=rs("cantidad")
	%>
	 <div class="page-header">
        <h1>Alertas</h1>
      </div>
       <div class="col-md-12">
	   <div class="col-xs-2"><div class="alert alert-success" role="alert"><strong>Clientes Registrados</strong><h3><a href="clientes.asp"><%=xclien%></a></h3></div></div>
	   <div class="col-xs-2"><div class="alert alert-info" role="alert"><strong>Gestiones Registradas</strong><h3><a href="gestion.asp?estado=0"><%=xgestion0%></a></h3></div></div>
	   <div class="col-xs-2"><div class="alert alert-warning" role="alert"><strong>Usuarios Registrados</strong><h3><a href="usuarios.asp?estado=0"><%=xusuario0%></a></h3></div></div>
	   <div class="col-xs-2"><div class="alert alert-info" role="alert"><strong>Mailing Recopilados</strong><h3><a href="mailing.asp?estado=0"><%=xmailing0%></a></h3></div></div>
	   <div class="col-xs-2"><div class="alert alert-danger" role="alert"><strong>Club Recopilados</strong><h3><a href="club.asp?estado=0"><%=xclub0%></a></h3></div></div>
	   </div>
	  <div class="col-md-12">
	  <div class="col-xs-6"><div class="alert alert-warning" role="alert"><strong>Gestiones Perdidas</strong><h3><a href="gestion.asp?estado=1"><%=xgestion1%></a></h3></div></div>
	  <div class="col-xs-6"><div class="alert alert-danger" role="alert"><strong>Gestiones Vencidas</strong><h3><a href="gestion.asp?estado=2"><%=xgestion2%></a></h3></div></div>
	  </div>
    <script src="js/ie10-viewport-bug-workaround.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.3/jquery.min.js"></script>
    <script>window.jQuery || document.write('<script src="js/vendor/jquery.min.js"><\/script>')</script>
    <script src="js/bootstrap.min.js"></script>
    <script src="js/ie10-viewport-bug-workaround.js"></script>
    <script src="js/offcanvas.js"></script>
</body>
</html>
