<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%Response.Buffer=True%> 
<!-- #include file="odbc.asp" -->
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
    <ul class="breadcrumb">
        <li class="breadcrumb-item"><a href="index.asp">Inicio</a></li>
        <li class="breadcrumb-item active">Clientes</li>
    </ul>
  <%
  if request.Form("inputapellido")<>"" then
	if request.Form("inputid")="0" then
	  	sql=" insert into erp_clien (razon,apellido,nombre,domicilio,localidad,provincia,pais,telefono1,telefono2,telefono3,mail1,mail2,mail3,documento) values ('" & request.Form("inputrazon") & "','" & request.Form("inputapellido") & "','" & request.Form("inputnombre") & "','" & request.Form("inputdomicilio") & "','" & request.Form("inputlocalidad") & "','" & request.Form("inputprovincia") & "','" & request.Form("inputpais") & "','" & request.Form("inputtelefono1") & "','" & request.Form("inputtelefono2") & "','" & request.Form("inputtelefono3") & "','" & request.Form("inputmail1") & "','" & request.Form("inputmail2") & "','" & request.Form("inputmail3") & "','" & request.Form("documento") & "')"
		set rs = oConn.Execute(sql)
    else
	  	sql=" update erp_clien set razon='" & request.Form("inputrazon") & "',	apellido='" & request.Form("inputapellido") & "',	nombre='" & request.Form("inputnombre") & "',	domicilio='" & request.Form("inputdomicilio") & "',	localidad='" & request.Form("inputlocalidad") & "',	provincia='" & request.Form("inputprovincia") & "',	pais='" & request.Form("inputpais") & "',	telefono1='" & request.Form("inputtelefono1") & "',		telefono2='" & request.Form("inputtelefono2") & "',	telefono3='" & request.Form("inputtelefono3") & "',		mail1='" & request.Form("inputmail1") & "',		mail2='" & request.Form("inputmail2") & "',	mail3='" & request.Form("inputmail3") & "',	documento='" & request.Form("inputmail3") & "' where idclien=" & request.Form("inputid") 
		set rs = oConn.Execute(sql)
	end if
	response.Redirect("index.asp")
  end if
  	xidclien=0
	xrazon=""
	xapellido=""
	xnombre=""
	xdomicilio=""
	xlocalidad=""
	xprovincia=""
	xpais=""
	xtelefono1=""
	xtelefono2=""
	xtelefono3=""
	xmail1=""
	xmail2=""
	xmail3=""
	xdocumento=""
  if request.QueryString("idclien")<>0 then
	sql=" select * from erp_clien  where idclien=" & request.QueryString("idclien")
	set rs = oConn.Execute(sql)
	IF not rs.eof  THEN
	xidclien=rs("idclien")
	xrazon=rs("razon")
	xapellido=rs("apellido")
	xnombre=rs("nombre")
	xdomicilio=rs("domicilio")
	xlocalidad=rs("localidad")
	xprovincia=rs("provincia")
	xpais=rs("pais")
	xtelefono1=rs("telefono1")
	xtelefono2=rs("telefono2")
	xtelefono3=rs("telefono3")
	xmail1=rs("mail1")
	xmail2=rs("mail2")
	xmail3=rs("mail3")
	xdocumento=rs("documento")
	end if
  end if
  %>
	<%
	call menu
	%>

      <div class="page-header">
        <h1><%if xidclien=0 then %>Alta de Clientes<%else%>Edita Clientes<%end if%></h1>
      </div>
	  <div class="container">

      <div class="row row-offcanvas row-offcanvas-right">

        <div class="col-xs-12 col-sm-9">
      <form  action="erp_abccli.asp" method="post">
        <div class="col-xs-4">
        <label >ID</label><input type="text" id="inputid" name="inputid" value="<%=xidclien%>" class="form-control" readonly="">
		</div>
		<div style="clear:left;"><hr></div>
        <div class="col-xs-4">
		<label >Razon Social</label><input type="text" id="inputrazon" name="inputrazon" value="<%=xrazon%>" class="form-control" placeholder="Razon Social" required="" autofocus=""></div>
        <div class="col-xs-4">
		<label >Apellido</label><input type="text" id="inputapellido" name="inputapellido" value="<%=xapellido%>" class="form-control" placeholder="Apellido" required="" autofocus=""></div>
        <div class="col-xs-4">
		<label >Nombre</label><input type="text" id="inputnombre" name="inputnombre" class="form-control" value="<%=xnombre%>" placeholder="Nombre" required="" autofocus=""></div>
        <div class="col-xs-4">
		<label >Domicilio y Nº</label><input type="text" id="inputdomicilio" name="inputdomicilio" class="form-control" value="<%=xdomicilio%>" placeholder="Domicilio y Nombre" required="" autofocus=""></div>
        <div class="col-xs-4">
		<label >Localidad</label><input type="text" id="inputlocalidad" name="inputlocalidad" class="form-control" value="<%=xlocalidad%>" placeholder="localidad" required="" autofocus=""></div>
        <div class="col-xs-4">
		<label >Provincia</label><input type="text" id="inputprovincia" name="inputprovincia" class="form-control" value="<%=xprovincia%>" placeholder="Provincia" required="" autofocus=""></div>
        <div class="col-xs-4">
		<label >Pais</label><input type="text" id="inputpais" name="inputpais" class="form-control" placeholder="Pais" value="<%=xpais%>" required="" autofocus=""></div>
        <div class="col-xs-4">
		<label >CUIT/DNI</label><input type="text" id="inputdocumento" name="inputdocumento" class="form-control" value="<%=xdocumento%>" placeholder="Documento" autofocus=""></div>
        <div class="col-xs-4">
		<label >Telefono 1</label><input type="text" id="inputtelefono1" name="inputtelefono1" class="form-control" value="<%=xtelefono1%>" placeholder="Telefono" required="" autofocus=""></div>
        <div class="col-xs-4">
		<label >Telefono 1</label><input type="text" id="inputtelefono2" name="inputtelefono2" class="form-control" value="<%=xtelefono2%>" placeholder="Telefono" autofocus=""></div>
        <div class="col-xs-4">
		<label >Telefono 1</label><input type="text" id="inputtelefono3" name="inputtelefono3" class="form-control" value="<%=xtelefono3%>" placeholder="Telefono" autofocus=""></div>
        <div class="col-xs-4">
		<label >Correo Electronico 1</label><input type="email" id="inputmail1" name="inputmail1" class="form-control" value="<%=xmail1%>" placeholder="Correo Electronico" required="" autofocus=""></div>
       <div class="col-xs-4">
	   <label >Correo Electronico 2</label><input type="email" id="inputmail2" name="inputmail2" class="form-control"  value="<%=xmail2%>" placeholder="Correo Electronico" autofocus=""></div>
        <div class="col-xs-4">
		<label >Correo Electronico 3</label><input type="email" id="inputmail3" name="inputmail3" class="form-control"  value="<%=xmail3%>" placeholder="Correo Electronico" autofocus=""></div>
		<div style="clear:left;"><hr></div>
		<div class="col-xs-4"><button class="btn btn-lg btn-primary btn-block" type="submit">Grabar</button></div>
      </form>
	</div></div></div>


    <script src="js/ie10-viewport-bug-workaround.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.3/jquery.min.js"></script>
    <script>window.jQuery || document.write('<script src="js/vendor/jquery.min.js"><\/script>')</script>
    <script src="js/bootstrap.min.js"></script>
    <script src="js/ie10-viewport-bug-workaround.js"></script>
    <script src="js/offcanvas.js"></script>

</body></html>