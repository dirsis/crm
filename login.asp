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
  <%
  if request.Form("inputEmail")<>"" then
  	sql=" select * from pass where mail='" & request.Form("inputEmail") & "' and clave = '" & request.Form("inputPassword") & "'"
	'response.Write(sql)
	set rs = oConn.Execute(sql)
	if not rs.eof then
		rs.movefirst
		session("apellido")=rs("apellido")
		session("mail")=rs("mail")
		session("nivel")=rs("nivel")
		Response.Cookies("mail")=rs("mail")
		Response.Cookies("idpass")=rs("idpass")
		session("idpass")=rs("idpass")
		call grabalog("acceso")
		response.Redirect("index.asp")
	end if
  else
	session("mail")=""
	Response.Cookies("mail")=""
	Response.Cookies("idpass")=""
	session("idpass")=""
  end if
  %>

    <div class="container">

      <form class="form-signin" action="login.asp" method="post">
        <h2 class="form-signin-heading">Ingreso Login:</h2>
        <label for="inputEmail" class="sr-only">Email address</label>
        <input type="text" id="inputEmail" name="inputEmail" class="form-control" placeholder="Correo Electronico" required="" autofocus="">
        <label for="inputPassword" class="sr-only">Password</label>
        <input type="password" name="inputPassword" id="inputPassword" class="form-control" placeholder="Clave" required="">
        <div class="checkbox">
          <label>
            <input type="checkbox" value="remember-me"> 
            Recordarme          </label>
        </div>
        <button class="btn btn-lg btn-primary btn-block" type="submit">Entrar</button>
        <p>&nbsp;</p>
        <p>No tiene cuenta?, <a href="abcpass.asp">registrese </a></p>
      </form>

    </div> <!-- /container -->


    <!-- IE10 viewport hack for Surface/desktop Windows 8 bug -->
    <script src="js/ie10-viewport-bug-workaround.js"></script>
  

</body></html>