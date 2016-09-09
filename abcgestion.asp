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
  </head>
  <body>
  <%
  if request.Form("inputasunto")<>"" then
		str1 = month(request.Form("inputfecha"))
		str2 = day(request.Form("inputfecha"))
		str3 = year(request.Form("inputfecha"))
		xfecha = str3 & "-" & str1 & "-" & str2
		str1 = month(request.Form("inputfcompro"))
		str2 = day(request.Form("inputfcompro"))
		str3 = year(request.Form("inputfcompro"))
		xfcompro = str3 & "-" & str1 & "-" & str2

			
	if request.Form("inputid")="0" then
	  	sql=" insert into gestion (idclien,asunto,tema,fecha,aquien,estado,fcompro,idpass,avisa) values ('" & request.Form("inputidclien") & "','" & request.Form("inputasunto") & "','" & request.Form("inputtema") & "','" & xfecha & "','" & request.Form("inputaquien") & "','" & request.Form("inputestado") & "','" & xfcompro & "'," & session("idpass") & "," & request.Form("inputavisa") & ")"
		'response.Write(sql)
		set rs = oConn.Execute(sql)
    else
	  	sql=" update gestion set idclien=" & request.Form("inputidclien") & ",	asunto='" & request.Form("inputasunto") & "',	tema='" & request.Form("inputtema") & "',	fecha='" & xfecha & "',	aquien='" & request.Form("inputaquien") & "',	estado='" & request.Form("inputestado") & "',fcompro='" & xfcompro & "',modifico=" & session("idpass") & ",avisa=" & request.Form("inputavisa") & " where idgestion=" & request.Form("inputid") 
		'response.Write(sql)	
		set rs = oConn.Execute(sql)
	end if
	response.Redirect("index.asp")
  else
	xidgestion=0
	xasunto=""
	xtema=""
	xfecha=date
	xaquien=""
	xestado=""
	xfcompro=date
	xavisa=0
	if request.QueryString("idclien")>0 then
		sql=" select * from clien where idclien=" & request.QueryString("idclien")
		set rs = oConn.Execute(sql)
		if not eof then
			xrazon=rs("razon")
			xapellido=rs("apellido")
			xnombre=rs("nombre")
			xidclien=rs("idclien")
		else
			response.Redirect("index.asp")
		end if  
	end if
  end if
  if request.QueryString("idgestion")<>0 then
	sql=" select g.idgestion,g.idclien,c.razon,c.apellido,c.nombre,g.asunto,g.tema,g.fecha,g.aquien,g.estado,g.fcompro,g.avisa from gestion g, clien c where c.idclien=g.idclien and g.idgestion=" & request.QueryString("idgestion")
	set rs = oConn.Execute(sql)
	IF not rs.eof  THEN
		xidgestion=rs("idgestion")
		xrazon=rs("razon")
		xapellido=rs("apellido")
		xnombre=rs("nombre")
		xasunto=rs("asunto")
		xtema=rs("tema")
		xfecha=rs("fecha")
		xaquien=rs("aquien")
		xestado=rs("estado")
		xfcompro=rs("fcompro")
		xidclien=rs("idclien")
		xavisa=rs("avisa")
	end if
  end if
  %>
	<%
	call menu
	%>
	<div class="page-header">
    	<h1>Edita Gestion</h1>
    </div>
    <div class="container">
      <div class="row row-offcanvas row-offcanvas-right">
        <div class="col-xs-12 col-sm-9">
      	<form  action="abcgestion.asp" method="post">
        <div class="col-xs-4">
        <label >ID</label><input type="text" id="inputid" name="inputid" value="<%=xidgestion%>" class="form-control" readonly="">
		</div>
		<div style="clear:left;"><hr></div>
        <div class="col-xs-4"><label >Id Cliente</label><input type="text" id="inputidclien" name="inputidclien" value="<%=xidclien%>" class="form-control" readonly ></div>
		<div class="col-xs-4"><label >Razon Social</label><input type="text" id="inputrazon" name="inputrazon" value="<%=xrazon%>" class="form-control" readonly></div>
        <div class="col-xs-4"><label >Apellido</label><input type="text" id="inputapellido" name="inputapellido" value="<%=xapellido%>" class="form-control" readonly></div>
        <div class="col-xs-4"><label >Nombre</label><input type="text" id="inputnombre" name="inputnombre" class="form-control" value="<%=xnombre%>" readonly></div>
		<div style="clear:left;"><hr></div>
		<div style="clear:left;"><hr></div>
		
        <div class="col-xs-12">
		<label >Asunto:</label><input type="text" id="inputasunto" name="inputasunto" class="form-control" value="<%=xasunto%>" placeholder="Asunto" required="" autofocus=""></div>
        <div class="col-xs-12">
		<label >Tema</label>
		<textarea rows="4" id="inputtema" name="inputtema"><%=xtema%></textarea>
		</div>
		<div style="clear:left;"><hr></div>
        <div class="col-xs-4">
		<label>Fecha Registro</label>
            <div class="hero-unit">
                <input  type="text" placeholder="click para ver calendario" id="inputfecha" name="inputfecha" value="<%=xfecha%>" class="form-control">
            </div>
		</div>
        <div class="col-xs-4">
			<label >Alarma?</label>
			<div class="input-group"><span class="input-group-addon"><input type="checkbox" name="inputavisa" id="inputavisa" class="form-control" value=1 <%if xavisa=1 then response.Write("checked") end if%>></span></div>							
			<label >A quien?</label>
			<select class="form-control input-sm" id="inputaquien" name="inputaquien">
			<%
				sql=" select * from pass "
				set rs = oConn.Execute(sql)
				do while not rs.eof
					%>
					<option value="<%=rs("idpass")%>" <%if rs("idpass")=xaquien then response.Write("selected") end if%> ><%=rs("apellido")%></option>
					<%
					rs.movenext
				loop
			%>
			</select>
		</div>
        <div class="col-xs-4">
		<label>Fecha Compromiso</label>
            <div class="hero-unit">
                <input  type="text" placeholder="click para ver calendario"  id="inputfcompro" name="inputfcompro"  value="<%=xfcompro%>" class="form-control" >
            </div>
		</div>
		<div style="clear:left;"><hr></div>
		<div class="col-xs-4">
		<button class="btn btn-lg btn-primary btn-block" type="submit">Grabar</button></div>
      	</form>
	</div></div></div>
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
        <script type="text/javascript">
            // When the document is ready
            $(document).ready(function () {
                
                $('#inputfcompro').datepicker({
                   language: "ar", 
				   format: "dd/mm/yyyy",
				   autoclose: true
                });  
            
            });
        </script>
</body></html>