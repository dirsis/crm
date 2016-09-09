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
  </head>
  <body>
  <%
  xestado=0
  if request.Form("btnterminar")="ok" then
	  	sql=" select * from club where tarjeta='" & request.Form("inputtarjeta") & "'"
		set rs = oConn.Execute(sql)
		if rs.eof then
			str1 = month(request.Form("inputfecha"))
			str2 = day(request.Form("inputfecha"))
			str3 = year(request.Form("inputfecha"))
			xfecha = str3 & "-" & str1 & "-" & str2
		  	sql=" insert into club (tarjeta,apellido,nombre,documento,fecnac,telefono,domicilio,localidad,provincia,correo,idpass) values ('" & request.Form("inputtarjeta") & "','" & request.Form("inputapellido") & "','" & request.Form("inputnombre") & "','" & request.Form("inputdocumento") & "','" & xfecha & "','" & request.Form("inputtelefono") & "','" & request.Form("inputdomicilio") & "','" & request.Form("inputlocalidad") & "','" & request.Form("inputprovincia") & "','" & request.Form("inputcorreo") & "'," & session("idpass") & ")"
			set rs = oConn.Execute(sql)
			xestado=2
		else
			xestado=1
		end if
  else
	xtarjeta=""
	xapellido=""
	xnombre=""
	xdocumento=""
	xfecha=""
	xdomicilio=""
	xlocalidad=""
	xprovincia=""
	xcorreo=""
	xtelefono=""
  end if
  %>
	<%
	call menu
	%>
    <div class="container">
      <div class="row row-offcanvas row-offcanvas-right">
      	<form  action="abcclub.asp" method="post">
		<div id="ErrorMsg" name="ErrorMsg" style="display:none">
			<input type="text" name="Mensaje" id="Mensaje" style="background-color:#B00002;  border-style: none;text-align: center;color: #fff;  width:100%;   font-size: 12px;
    font-family: verdana;" readonly="">
		</div>
        <div class="col-xs-12 col-sm-9" id="primero" name="primero">
        <h2 class="form-signin-heading">Registro de Tarjeta Club</h2>
        <div class="col-xs-4"><label id="linputtarjeta">Numero de Tarjeta Club (*)</label><input type="text" id="inputtarjeta"  name="inputtarjeta" value="<%=xidclien%>" class="form-control" required="" autofocus=""></div>
        <div class="col-xs-4"><%if xestado=1 then%><p style="color:#FF0000; font-size:16px;"><B>TARJETA YA REGISTRADA!!</B></p><%end if%></div>
		<div style="clear:left;"><hr></div>
        <div class="col-xs-4"><label id="linputapellido" >Apellido (*)</label><input type="text" id="inputapellido" name="inputapellido" value="<%=xapellido%>" class="form-control" required="" autofocus=""></div>
        <div class="col-xs-4"><label id="linputnombre" >Nombre (*)</label><input type="text" id="inputnombre" name="inputnombre" class="form-control" value="<%=xnombre%>" required="" autofocus=""></div>
		<div style="clear:left;"><hr></div>
		<div class="col-xs-4"><label  id="linputdocumento">Documento (*)</label><input type="text" id="inputdocumento" name="inputdocumento" value="" class="form-control" required="" autofocus=""></div>
        <div class="col-xs-4"><label  id="linputfecha">Fecha de Nacimiento dd/mm/aaaa(*)</label><input  type="text" id="inputfecha" name="inputfecha" value="<%=xfecha%>" class="form-control" required="" autofocus=""  onkeyup="myFunction()" maxlength="10">
		<script>
		function myFunction() {
		    var e = document.getElementById("inputfecha");
			if (e.keyCode == 8){  //BackSpace
			}
			else {
				if (e.value.length === 2) {
					e.value = e.value + "/";
				}
				if (e.value.length === 5) {
					e.value = e.value + "/";
				}
			}	
		}
		</script>
		</div>
		<div style="clear:left;"><hr></div>
		<div class="col-xs-4">
		<a href="#" id="btnIngresar" name="btnIngresar"  class="btn btn-lg btn-primary btn-block" onClick="Visibility2()" tabindex="10">Continuar</a> 
		</div>
		<div style="clear:left;"><hr></div>
        <div class="col-xs-12 col-sm-9" id="segundo" name="segundo" style="display:none">
        <div class="col-xs-6"><label id="linputdomicilio" >Domicilo y Numero (*)</label><input type="text" id="inputdomicilio" name="inputdomicilio" value="<%=xdomicilio%>" class="form-control" required="" autofocus=""></div>
        <div class="col-xs-6"><label id="linputlocalidad" >Localidad (*)</label><input type="text" id="inputlocalidad" name="inputlocalidad" value="<%=xlocalidad%>" class="form-control" required="" autofocus=""></div>
        <div class="col-xs-6"><label id="linputprovincia" >Provincia (*)</label><input type="text" id="inputprovincia" name="inputprovincia" value="<%=xprovincia%>" class="form-control" required="" autofocus=""></div>
		<div style="clear:left;"><hr></div>
        <div class="col-xs-6"><label id="linputtelefono" >Telefono (*)</label><input type="text" id="inputtelefono" name="inputtelefono" class="form-control" value="<%=xtelefono%>" required="" autofocus=""></div>
		<div class="col-xs-8"><label id="linputcorreo" >Correo Electronico (*)</label><input type="email" id="inputcorreo" name="inputcorreo" value="<%=xcorreo%>" class="form-control" required="" autofocus=""></div>
		<div style="clear:left;"><hr></div>
		<div class="campos">
				<table name="tablaCheckBox" id="tablaCheckBox" style="display:block">
				<tbody><tr>
				<td colspan="1" align="left"><input type="checkbox" style="width:30px; height:15px;" checked="" class="input-terminos" name="aceptonovedades" id="aceptonovedades" tabindex="29"></td>
				<td><label>Estoy de acuerdo en recibir novedades y promociones vía e-mail</label></td>
				</tr>
				<tr style="justify-content:flex-start;">
				<td colspan="1" align="left"><input type="checkbox" style="width:30px; height:15px;" class="input-terminos" name="aceptoterminos" id="aceptoterminos" tabindex="30"></td>
				<td><label style="cursor:pointer" onClick="Visibility()"><u>Estoy de acuerdo con los Términos y Condiciones. Ver términos</u></label></td>
				</tr>
				</tbody></table>
		</div>
		<div id="textoTerminos" name="textoTerminos" style="display:none; background-color:#FFFFFF; border:solid;">
		<b>Tarjeta CLUB</b><br><b>TÉRMINOS Y CONDICIONES DEL PROGRAMA</b>
		<ol>
		<li><strong>Definiciones</strong>
		<ul>
		<li>La confirmación del presente formulario da por aceptado estos terminos y condiciones.</li>
		<li>La información aqui registrada solo va ser usada por nuestra empresa a fines estadisticos o promocionales.</li>
		<li>Nos comprometemos a respetar su información y valorarla como privada y bajo relacion estrictamente comercial.</li>
		</ul>
		</li>
		</ol>
		<a href="#" id="btnOcultarTerminos" name="btnOcultarTerminos" class="ingresar" onClick="Visibility()">Ocultar</a>
		</div>
		<div class="col-xs-4">
				<a href="#" id="btncontinuar" name="btncontinuar"  class="btn btn-lg btn-primary btn-block" onClick="Visibility3()" tabindex="10">Continuar</a> 
				<button class="btn btn-lg btn-primary btn-block" type="submit" id="btnterminar" name="btnterminar" style="display:none;" value="ok">Terminar</button>
		</div>
		</div>		
      	</form>
	</div>
	<script>
			function Visibility() {
				var var_text = document.getElementById('textoTerminos');   

				if(var_text.style.display == 'block') {
					var_text.style.display = 'none';
				} 
				else {
					var_text.style.display = 'block';
				}
			}	
			function Visibility2() {
				var var_text = document.getElementById('segundo'),
					var_inputtarjeta = document.getElementById('inputtarjeta'),
					var_linputtarjeta = document.getElementById('linputtarjeta'),
					var_inputapellido = document.getElementById('inputapellido'),
					var_linputapellido = document.getElementById('linputapellido'),
					var_inputnombre = document.getElementById('inputnombre'),
					var_linputnombre = document.getElementById('linputnombre'),
					var_inputdocumento = document.getElementById('inputdocumento'),
					var_linputdocumento = document.getElementById('linputdocumento'),
					var_inputfecha = document.getElementById('inputfecha'),
					var_linputfecha = document.getElementById('linputfecha'),
					var_btnIngresar = document.getElementById('btnIngresar');

				if (validateNumber(var_inputtarjeta.value)) {
					var_linputtarjeta.style.color= '#00CC00';
					if(validateText(var_inputapellido.value)) {
						var_linputapellido.style.color= '#00CC00';
						if(validateText(var_inputnombre.value)) {
							var_linputnombre.style.color= '#00CC00';
							if(validateNumber(var_inputdocumento.value)) {
								var_linputdocumento.style.color= '#00CC00';
								if(validateDate(var_inputfecha.value)) {
									var_linputfecha.style.color= '#00CC00';
									var_btnIngresar.style.display = 'none';
									var_text.style.display = 'block';
								}
								else {
									var_linputfecha.style.color= '#FF3300';
									var_inputfecha.placeholder='10/10/1970';
									var_inputfecha.value='';
								}
							}
							else {
								var_linputdocumento.style.color= '#FF3300';
								var_inputdocumento.placeholder='Ingresar DNI/CUIL/CUIT';
								var_inputdocumento.value='';
							}
						}
						else {
							var_linputnombre.style.color= '#FF3300';
							var_inputnombre.placeholder='Ingresar Nombre...';
							var_inputnombre.value='';
						}
					}
					else {
						var_linputapellido.style.color= '#FF3300';
						var_inputapellido.placeholder='Ingresar Apellido...';
						var_inputapellido.value='';
					}
				}
				else {
					var_linputtarjeta.style.color= '#FF3300';
					var_inputtarjeta.placeholder='Ej: 015259';
					var_inputtarjeta.value='';

				}
			}	

			function Visibility3() {
				var var_linputcorreo = document.getElementById('linputcorreo');
				var var_inputcorreo = document.getElementById('inputcorreo');
				var var_btncontinuar = document.getElementById('btncontinuar');
				var var_btnterminar = document.getElementById('btnterminar');
				if (validateCorreo(var_inputcorreo.value)) {
					var_linputcorreo.style.color= '#00CC00';
					var_btncontinuar.style.display = 'none';
					var_btnterminar.style.display = 'block';
				}
				else {
					var_linputcorreo.style.color= '#FF3300';
					var_inputcorreo.placeholder='Ingresar correo...';
					var_inputcorreo.value='';
				}
			}	
			
	</script>
	
    <script src="js/ie10-viewport-bug-workaround.js"></script>
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/1.11.3/jquery.min.js"></script>
    <script>window.jQuery || document.write('<script src="js/vendor/jquery.min.js"><\/script>')</script>
    <script src="js/bootstrap.min.js"></script>
    <script src="js/ie10-viewport-bug-workaround.js"></script>
    <script src="js/offcanvas.js"></script>
        <!-- Load jQuery and bootstrap datepicker scripts -->
        <script src="js/jquery.min.js"></script>

        <script src="js/bootstrap-datepicker.js"></script>
        <script src="js/locales/bootstrap-datepicker.es.min.js"></script>
        <script type="text/javascript">
            // When the document is ready
            $(document).ready(function () {
                
                $('#inputfecha').datepicker({
                    language: "ar",
					format: "dd/mm/yyyy",
					autoclose: true
                });  
            
            });
        </script>
</body></html>