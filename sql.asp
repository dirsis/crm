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
    <meta name="description" content="">
    <meta name="author" content="">
    <link rel="icon" href="http://getbootstrap.com/favicon.ico">
    <title>SQL</title>
    <link href="css/bootstrap.min.css" rel="stylesheet">
    <link href="css/ie10-viewport-bug-workaround.css" rel="stylesheet">
    <link href="css/signin.css" rel="stylesheet">
    <script src="js/ie-emulation-modes-warning.js"></script>
  </head>
  <body>
  <%
  if request.Form("inputbusca")<>"" then
		sql=request.Form("inputbusca")
		set rs = oConn.Execute(sql)
  end if
  %>
	<%
	call menu
	%>
      <div class="page-header">
        <h1>Recopilados</h1>
      </div>
      <div class="row">
        <div class="col-md-12">
       		<form class="form-signin" action="sql.asp" method="post">
       		<label >Buscador</label>
       		<textarea name="inputbusca" cols="200" rows="6" id="inputbusca" required="" autofocus=""></textarea>
			<button class="btn btn-lg btn-primary btn-block" type="submit">Buscar</button>
			</form>
		</div>
	</div>
    <script src="js/ie10-viewport-bug-workaround.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.3/jquery.min.js"></script>
    <script>window.jQuery || document.write('<script src="js/vendor/jquery.min.js"><\/script>')</script>
    <script src="js/bootstrap.min.js"></script>
    <script src="js/ie10-viewport-bug-workaround.js"></script>
    <script src="js/offcanvas.js"></script>
	</body>
</html>