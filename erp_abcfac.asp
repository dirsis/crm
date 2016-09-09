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
  if request.Form("inputmotivo")<>"" then
		str1 = month(request.Form("inputfecha"))
		str2 = day(request.Form("inputfecha"))
		str3 = year(request.Form("inputfecha"))
		xfecha = str3 & "-" & str1 & "-" & str2
		xnumero=0
	if request.Form("inputcomprob")="0" then
	  	sql=" insert into erp_comprob (idclien,motivo,total,fecha,numero,idpass) values ('" & request.Form("inputidclien") & "','" & request.Form("inputmotivo") & "'," & request.Form("inputtotal") & ",'" & xfecha & "'," & xnumero & "," & session("idpass") & ")"
		set rs = oConn.Execute(sql)
		xclien=request.Form("inputidclien")
    else
	  	sql=" update erp_comprob set idclien=" & request.Form("inputidclien") & ",	motivo='" & request.Form("inputmotivo") & "',	total=" & request.Form("inputtotal") & ",	fecha='" & xfecha & "', idpass= " & session("idpass") & " where idcomprob=" & request.Form("inputcomprob") 
		'response.Write(sql)	
		set rs = oConn.Execute(sql)
		xclien=request.Form("inputidclien")
	end if
	response.Redirect("index.asp")
  else
	xidcomprob=0
	xmotivo=""
	xtotal=0
	xfecha=date
	if request.QueryString("idclien")>0 then
		sql=" select * from erp_clien where idclien=" & request.QueryString("idclien")
		set rs = oConn.Execute(sql)
		if not eof then
			xidclien=rs("idclien")
		else
			response.Redirect("index.asp")
		end if  
	end if
  end if
  if request.QueryString("idcomprob")<>0 then
	sql=" select g.idcomprob,g.idclien,c.razon,c.apellido,c.nombre,g.motivo,g.total,g.fecha,g.numero from erp_comprob g, erp_clien c where c.idclien=g.idclien and g.idcomprob=" & request.QueryString("idcomprob")
	set rs = oConn.Execute(sql)
	IF not rs.eof  THEN
		xidcomprob=rs("idcomprob")
		xrazon=rs("razon")
		xapellido=rs("apellido")
		xnombre=rs("nombre")
		
		xmotivo=rs("motivo")
		xtotal=rs("total")
		xfecha=rs("fecha")
		xidclien=rs("idclien")
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
      	<form  action="erp_abcfac.asp" method="post">
        <div class="col-xs-12"><div class="col-xs-4"><label>ID</label><input type="text" id="inputcomprob" name="inputcomprob" value="<%=xidcomprob%>" class="form-control" readonly=""></div></div>
        <div class="col-xs-12"><div class="col-xs-4"><label>Fecha Registro</label><input  type="text" id="inputfecha" name="inputfecha" value="<%=xfecha%>" class="form-control"></div></div>
		
		<div class="col-xs-12"><div class="col-xs-6"><label>Id Cliente</label><select class="form-control input-sm" id="inputidclien" name="inputidclien">
			<%
				sql=" select * from erp_clien order by apellido"
				set rs = oConn.Execute(sql)
				do while not rs.eof
					%>
					<option value="<%=rs("idclien")%>" <%if rs("idclien")=xidclien then response.Write("selected") end if%> ><%=rs("apellido")%> (<%=rs("nombre")%>)</option>
					<%
					rs.movenext
				loop
			%>
			</select>
		</div></div>
		<div class="col-xs-12"><div class="col-xs-6"><label >Motivo:</label></div><div class="col-xs-3"><label >$ Importe:</label></div></div>
        <div class="col-xs-12">
        <div class="col-xs-6"><input type="text" id="inputmotivo" name="inputmotivo" class="form-control" value="<%=xmotivo%>" placeholder="Motivo..." required="" autofocus=""></div>
        <div class="col-xs-3"><input type="text" id="inputtotal" name="inputtotal" class="form-control" value="<%=xtotal%>" placeholder="0" required="" autofocus=""></div>
		</div>
		<div class="col-xs-12"><div class="col-xs-3">
		<button class="btn btn-lg btn-primary btn-block" type="submit">Grabar</button>
		</div>
		</div>
      	</form>
	</div></div></div>
<%if xidclien<>"" then%>
	<div class="col-md-12">
          <table class="table" width="100%">
            <thead>
              <tr>
	            <th>Fecha</th>
                <th>motivo</th>
                <th>Total</th>
                <th>Acumulado</th>
                <th>Nota</th>
                <th>Registro</th>
              </tr>
            </thead>
            <tbody>
			  <%
					sql=" select g.idcomprob,g.motivo,g.total,g.fecha,g.nota,p.apellido as registro,c.apellido from erp_comprob g, erp_clien c, pass p where g.idclien=c.idclien and g.idpass=p.idpass "
					sql=sql & " and g.idclien='" & xidclien & "' "
					sql=sql & " order by g.fecha"
					set rs = oConn.Execute(sql)
					IF not rs.eof  THEN
					rs.movefirst
					xacum=0
					do while not rs.eof 
					xacum=xacum+rs("total")
					%>
	                <tr>
                	<td><%=rs("fecha")%></td>
                	<td><%=rs("motivo")%></td>
                	<td><%=rs("total")%></td>
                	<td><%=xacum%></td>
                	<td><%=rs("nota")%></td>
                	<td><%=rs("registro")%></td>
              	    </tr>
					<%
					rs.movenext
					loop
					end if
				%>
            </tbody>
          </table>
        </div>
<%end if%>			
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
	</body>
</html>