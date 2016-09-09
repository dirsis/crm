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
    <link rel="stylesheet" href="css/bootstrap-datetimepicker">
    <link rel="stylesheet" href="css/bootstrap.css">
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
	  <script src="tinymce/tinymce.min.js"></script>
	  <script>tinymce.init({ selector:'textarea' });</script>	
	  <script>
			function validateCorreo(value) {
				expr =  /^((([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+(\.([a-z]|\d|[!#\$%&'\*\+\-\/=\?\^_`{\|}~]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])+)*)|((\x22)((((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(([\x01-\x08\x0b\x0c\x0e-\x1f\x7f]|\x21|[\x23-\x5b]|[\x5d-\x7e]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(\\([\x01-\x09\x0b\x0c\x0d-\x7f]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF]))))*(((\x20|\x09)*(\x0d\x0a))?(\x20|\x09)+)?(\x22)))@((([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|\d|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))\.)+(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])|(([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])([a-z]|\d|-|\.|_|~|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])*([a-z]|[\u00A0-\uD7FF\uF900-\uFDCF\uFDF0-\uFFEF])))$/i;
				return expr.test(value);
			}
			
			function validateNumber(value) {
				expr =  /^\d+$/;
				return expr.test(value);
			}
			
			function validateText(value) {
				expr =  /^[A-Za-z ñáéíóúÑÁÉÍÓÚ\_\-\.\s\xF1\xD1]+$/;
				return expr.test(value);
			}
			
			function validateDate(value) {
				var arrayFecha = value.split("/");
				
				var inputValues = arrayFecha[0] + ' ' + arrayFecha[1] + ' ' + arrayFecha[2];

				var d = new Date();
				var n = d.getHours();
				var m = d.getMinutes();
				var p = d.getSeconds();

				var date = arrayFecha[0];
				var month = arrayFecha[1];
				var year = arrayFecha[2];

				var dateCheck = /^(0?[1-9]|[12][0-9]|3[01])$/;
				var monthCheck = /^(0[1-9]|1[0-2])$/;
				var yearCheck = /^\d{4}$/; 

				if (month.match(monthCheck) && date.match(dateCheck) && year.match(yearCheck)) {

					var ListofDays = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31];

					if (month == 1 || month > 2) {
						if (date > ListofDays[month - 1]) {
							return false;
						}
					}

					if (month == 2) {
						var leapYear = false;
						if ((!(year % 4) && year % 100) || !(year % 400)) {
							leapYear = true;
						}

						if ((leapYear == false) && (date >= 29)) {
							return false;
						}

						if ((leapYear == true) && (date > 29)) {
							return false;
						}

					}
					var flag = 1
				}

				else {
					return false;


				}
				if (flag == 1) {
					return true;
				}
			}
  
	  </script>
		<script type="text/javascript">
        	function mayuscula()
			{
			    var x1 = document.getElementById("inputapellido"),
				x2 = document.getElementById("inputnombre"),
				x3 = document.getElementById("inputlocalidad"),
				x4 = document.getElementById("inputtelefono");
				
    			x1.value = x1.value.toUpperCase();
	    		x2.value = x2.value.toUpperCase();
	    		x3.value = x3.value.toUpperCase();
	    		x4.value = x4.value.toUpperCase();
    	    }
	    </script>
  </head>
  <body>
  <%
  if request.Form("btnterminar")="ok" then
  	if request.form("inputidmailing")<>0 then
  		sql=" update mailing  set apellido='" & request.Form("inputapellido") & "',	nombre='" & request.Form("inputnombre") & "',fecnac='" & request.Form("inputfecnac") & "',localidad='" & request.Form("inputlocalidad") & "',telefono='" & request.Form("inputtelefono") & "',correo='" & request.Form("inputcorreo") & "',idpass=" & session("idpass") & ",motivo='" & request.Form("inputmotivo") & "',zona='" & request.Form("inputmotivo") & "' where idmailing= " & request.form("inputidmailing")
	else
	  	sql=" insert into mailing (apellido,nombre,fecnac,localidad,telefono,correo,idpass,motivo,zona) values ('" & request.Form("inputapellido") & "','" & request.Form("inputnombre") & "','" & request.Form("inputfecnac") & "','" & request.Form("inputlocalidad") & "','" & request.Form("inputtelefono") & "','" & request.Form("inputcorreo") & "'," & session("idpass") & ",'" & request.Form("inputmotivo") & "','" & request.Form("inputzona") & "')"
	end if
	set rs = oConn.Execute(sql)
	msg="Registrado '" & request.form("inputapellido") &"," & request.form("inputnombre") & "'"
  else
  	xidmailing=0
	xapellido=""
	xnombre=""
	xfecnac=""
	xlocalidad=""
	xtelefono=""
	xcorreo=""
	xmotivo=""
  	if request.QueryString("idmailing")<>0 then
	  	sql=" select * from mailing  where idmailing= " & request.QueryString("idmailing")
		set rs = oConn.Execute(sql)
		if not rs.eof then
			xidmailing=rs("idmailing")
			xapellido=rs("apellido")
			xnombre=rs("nombre")
			xfecnac=rs("fecnac")
			xzona=rs("zona")
			xlocalidad=rs("localidad")
			xtelefono=rs("telefono")
			xcorreo=rs("correo")
			xmotivo=rs("motivo")
		end if
	end if
  end if
	call menu
	%>
	<div class="page-header">
    	<h1>Registro de Informacion de General</h1>
    </div>
	<div class="container">
      <div class="row row-offcanvas row-offcanvas-right">
      	<form  action="abcmailing.asp" method="post">
        <div class="col-xs-4">
        <label >ID</label><input type="text" id="inputidmailing" name="inputidmailing" value="<%=xidmailing%>" class="form-control" readonly="">
		</div>
		<div style="clear:left;"><hr></div>
        <div class="col-xs-12 col-sm-9" id="primero" name="primero">
        <div class="col-xs-4"><label id="linputapellido" >Apellido (*)</label>
		<input type="text" id="inputapellido" name="inputapellido" value="<%=xapellido%>" class="form-control" onKeyUp="mayuscula()"  required="" autofocus=""></div>
        <div class="col-xs-4"><label id="linputnombre" >Nombre (*)</label>
		<input type="text" id="inputnombre" name="inputnombre" class="form-control" value="<%=xnombre%>" onKeyUp="mayuscula()" required="" autofocus=""></div>
		<div style="clear:left;"><hr></div>
        <div class="col-xs-4"><label  id="linputfecha">Fecha de Nacimiento (dd/mm , dd/mm, etc...)(*)</label>
		<input  type="text" id="inputfecnac" name="inputfecnac" value="<%=xfecnac%>" class="form-control" autofocus="" >
		</div>
        <div class="col-xs-4"><label id="linputlocalidad" >Localidad (*)</label><input type="text" id="inputlocalidad" name="inputlocalidad" onKeyUp="mayuscula()" value="<%=xlocalidad%>" class="form-control" required="" autofocus=""></div>
        <div class="col-xs-6"><label id="linputtelefono" >Telefono (*)</label><input type="text" id="inputtelefono" name="inputtelefono" onKeyUp="mayuscula()" class="form-control" value="<%=xtelefono%>" required="" autofocus=""></div>
		<div class="col-xs-8"><label id="linputcorreo" >Correo Electronico (*)</label><input type="text" id="inputcorreo" name="inputcorreo" value="<%=xcorreo%>" class="form-control" required="" autofocus=""></div>
		<div class="col-xs-8"><label id="linputmotivo" >Motivo</label><input type="text" id="inputmotivo" name="inputmotivo" value="<%=xmotivo%>" class="form-control" autofocus=""></div>
		<div class="col-xs-8"><label id="linputzona" >Zona</label><input type="text" id="inputzona" name="inputzona" value="<%=xzona%>" class="form-control" autofocus=""></div>
		<div style="clear:left;"><hr></div>
		<div class="col-xs-4">
				<button class="btn btn-lg btn-primary btn-block" type="submit" id="btnterminar" name="btnterminar" value="ok">Terminar</button>
		</div>
		<div class="col-xs-4">
				<span class="label label-success"><%=msg%></span>
		</div>
		</div>		
      	</form>
	</div>
    <script src="js/ie10-viewport-bug-workaround.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.3/jquery.min.js"></script>
    <script>window.jQuery || document.write('<script src="js/vendor/jquery.min.js"><\/script>')</script>
    <script src="js/bootstrap.min.js"></script>
    <script src="js/ie10-viewport-bug-workaround.js"></script>
    <script src="js/offcanvas.js"></script>
	<script src="js/jquery.min.js"></script>
	<script src="js/bootstrap-datepicker.js"></script>
	<script src="js/locales/bootstrap-datepicker.es.min.js"></script>
</body></html>
