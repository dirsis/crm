<%
session("activo")="S"
session("path")="web/"
session("recargo")=0
session("CANTIDADITEMS")=9
session("COLUMNAS")=3
session("BASEDATO")="toys_ecommerce"
session("Dolar")=8
session("DESCUENTO")=0
'session("codigo")="9999"
'
nombre_archivo = "odbc.txt" 
Set fs = Server.CreateObject("Scripting.FileSystemObject") 
Set archivo = fs.OpenTextFile(Server.MapPath(nombre_archivo), 1) 
do while not archivo.AtEndOfStream  
    xlinea=Split(archivo.ReadLine, ";")
	if xlinea(0)="database" then
		session("BASEDATO")=xlinea(1)
	end if
	if xlinea(0)="themes" then
		session("imagenes")=xlinea(1)
	end if
	if xlinea(0)="fin" then
		exit do
	end if
loop 
archivo.Close 
Set archivo = Nothing 
Set fs = Nothing
'
Conecta="PROVIDER=SQLOLEDB;DATA SOURCE=localhost;UID=sa;DATABASE=" & session("BASEDATO")
'para base en hosting dattatec
Conecta="PROVIDER=SQLOLEDB;data source=localhost;integrated security=SSPI;initial catalog=" & session("BASEDATO")
'-conecta="Provider=SQLNCLI10;Server=localhost;Database=toys_clinic; Trusted_Connection=yes;"
'
set oConn = Server.CreateObject("ADODB.Connection")
oConn.open(Conecta)


Sub MailSender(desdecorreo,asunto,tema,paracorreo)
	usuario=txttxt("smtp")
	passwd=txttxt("clave")
	correoorg=txttxt("smtp")
	if paracorreo="" then
		paracorreo=correoorg
	end if
	If Not paracorreo = "" Then
' Se crean los objetos necesarios para el envío del correo
		Set oMail = Server.CreateObject("CDO.Message") 
		Set iConf = Server.CreateObject("CDO.Configuration") 
		Set Flds = iConf.Fields 
		
		' Se configuran los parametros necesarios para el envío
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 1 
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "http://127.0.0.1" 
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10 
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
		' Se completan los datos del usuario y la contraseña necesarios para el envio
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = usuario 'usuario smtp
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = passwd  'password para STMP
		iConf.Fields.Update
		' Se asignan las propiedades de configuración al objeto
		Set oMail.Configuration = iConf 
	
		' Destinatario del correo
		oMail.To = paracorreo
		' Remitente del correo
		oMail.From = correoorg ' "noreply@ferozowindows.com.ar"
		' Subject o asunto
		oMail.Subject = asunto
		' Cuerpo del mensaje
		oMail.HTMLBody = tema
		' Se envía el correo
		oMail.Send
		' Se destruyen los objetos
		Set iConf = Nothing 
		Set Flds = Nothing
	end if
end sub

Sub MailSendergmail(asunto,tema,paracorreo)
	desdecorreo="dirsis@gmail.com"
	if paracorreo="" then
		paracorreo="dirsis@gmail.com"
	end if
	If Not paracorreo = "" Then
' Se crean los objetos necesarios para el envío del correo
		Set oMail = Server.CreateObject("CDO.Message") 
		Set iConf = Server.CreateObject("CDO.Configuration") 
		Set Flds = iConf.Fields 
		
		' Se configuran los parametros necesarios para el envío
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2 
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "smtp.gmail.com"
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 10 
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 465
		' Se completan los datos del usuario y la contraseña necesarios para el envio
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendusername") = "dirsis@gmail.com"
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/sendpassword") = "Ravera951753"
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = 1
		iConf.Fields.Item("http://schemas.microsoft.com/cdo/configuration/smtpusessl") = 1

		iConf.Fields.Update
		' Se asignan las propiedades de configuración al objeto
		Set oMail.Configuration = iConf 
	
		' Destinatario del correo
		oMail.To = paracorreo
		' Remitente del correo
		oMail.From = "dirsis@gmail.com"
		' Subject o asunto
		oMail.Subject = asunto
		' Cuerpo del mensaje
		oMail.HTMLBody = tema
		' Se envía el correo
		oMail.Send
		' Se destruyen los objetos
		Set iConf = Nothing 
		Set Flds = Nothing
	end if
end sub


Function numerador(campo)
sql="select * from contador" 
set rs = oConn.Execute(sql)
rs.movefirst
if not rs.eof then
	xcampo=rs(campo)
	sql="update contador set " & campo & "=" & campo & "+1"
	set rs = oConn.Execute(sql)
	numerador=xcampo+1
end if
end Function

function menu
%>
	<!--
    <nav class="navbar navbar-fixed-top navbar-inverse">
      <div class="container">
        <div class="navbar-header">
          <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#navbar" aria-expanded="false" aria-controls="navbar">
            <span class="sr-only">Toggle navigation</span>
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
          </button>
          <a class="navbar-brand" href="#">CRM</a>
        </div>
        <div id="navbar" class="collapse navbar-collapse">
          <ul class="nav navbar-nav">
		  	<%if session("idpass") <>"" and session("nivel")<2 then%>
			<li><a href="abccli.asp">Nuevo<br />Cliente</a></li>
			<li><a href="gestion.asp">Gestiones</a></li>
			<li><a href="clientes.asp">Clientes</a></li>
			<li><a href="pass.asp">Usuarios</a></li>
			<li><a href="mailing.asp">Recopilados</a></li>
			<%end if%>
			<%if session("idpass") <>"" then%>
			<li><a href="abcmailing.asp">Nuevo<br />Recopilado</a></li>
            <li><a href="closelogin.asp">Cierra Sesion</a></li>
			<%end if%>
            <li><a href="index.asp">Inicio</a></li>
            <li style="color:#FFFF33">Usuario:(<%=session("idpass")%>) <%=session("apellido")%></li>
          </ul>
        </div>
      </div>
    </nav>
	-->
	<nav class="navbar navbar-fixed-top navbar-inverse">
  <div class="container-fluid">
    <div class="navbar-header">
          <button type="button" class="navbar-toggle collapsed" data-toggle="collapse" data-target="#navbar" aria-expanded="false" aria-controls="navbar">
            <span class="sr-only">Nav</span>
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
            <span class="icon-bar"></span>
          </button>
	<a class="navbar-brand" href="#">Distribuidora Pluto's</a>
    </div>
	<div id="navbar" class="collapse navbar-collapse">
    <ul class="nav navbar-nav">
      <li class="active"><a href="index.asp">Inicio</a></li>
	  <%if session("idpass") <>"" and session("nivel")<=2 then%>
      <li class="dropdown" >
        <a class="dropdown-toggle" data-toggle="dropdown" href="#">Gestion
        <span class="caret"></span></a>
        <ul class="dropdown-menu">
          <li><a role="menuitem" tabindex="-1" href="abccli.asp">Nuevo Franquiciado</a></li>
          <li><a href="clientes.asp">Lista Franquiciados</a></li>
		   <li role="presentation" class="divider"></li>
          <li><a href="abcgestion.asp">Nueva Gestion</a></li> 
          <li><a href="gestion.asp">Lista Gestiones</a></li> 
        </ul>
      </li>
	  <%end if%>
      <li class="dropdown">
        <a class="dropdown-toggle" data-toggle="dropdown" href="#">Recopilados
        <span class="caret"></span></a>
        <ul class="dropdown-menu">
          <li><a href="abcmailing.asp">Nuevo</a></li>
		  <%if session("idpass") <>"" and session("nivel")<=2 then%>
          <li><a href="mailing.asp">Lista</a></li>
		  <%end if%>
        </ul>
      </li>
      <li class="dropdown">
        <a class="dropdown-toggle" data-toggle="dropdown" href="#">Club
        <span class="caret"></span></a>
        <ul class="dropdown-menu">
          <li><a href="abcclub.asp">Alta</a></li>
		  <%if session("idpass") <>"" and session("nivel")<=2 then%>
          <li><a href="club.asp">Lista</a></li>
		  <%end if%>
        </ul>
      </li>
	  <%if session("idpass") <>"" and session("nivel")<=1 then%>
      <li class="dropdown">
        <a class="dropdown-toggle" data-toggle="dropdown" href="#">Facturacion
        <span class="caret"></span></a>
        <ul class="dropdown-menu">
          <li><a href="erp_abccli.asp">Cliente Nuevo</a></li>
          <li><a href="erp_clientes.asp">Lista Clientes</a></li>
          <li><a href="erp_abcfac.asp">Facturar</a></li>
          <li><a href="erp_historia.asp">Historia</a></li>
        </ul>
      </li>
	  <%end if%>
	  <%if session("idpass") <>"" and session("nivel")<=1 then%>
	    <li class="dropdown">
        <a class="dropdown-toggle" data-toggle="dropdown" href="#">Sistemas
        <span class="caret"></span></a>
        <ul class="dropdown-menu">
          <li><a href="abcpass.asp">Alta</a></li>
		  <%if session("idpass") <>"" and session("nivel")=1 then%>
          <li><a href="pass.asp">Lista</a></li>
		  <%end if%>
        </ul>
      </li>
	  <%end if%>
	 <li class="dropdown">
        <a class="dropdown-toggle" data-toggle="dropdown" href="#">Perfil
        <span class="caret"></span></a>
        <ul class="dropdown-menu">
          <li class="active">&nbsp;Usuario:<span class="badge"><%=session("idpass")%></span><%=session("apellido")%></li>
          <li class="active">&nbsp;Nivel:<span class="badge"><%=session("nivel")%></span></li>		  
          <li><a href="closelogin.asp">Cerrar Session</a></li>
        </ul>
      </li>
    </ul>
  </div>
  </div>
</nav>
<%
end function

function txttxt(campo)
nombre_archivo = "odbc.txt" 
Set fs = Server.CreateObject("Scripting.FileSystemObject") 
Set archivo = fs.OpenTextFile(Server.MapPath(nombre_archivo), 1) 
do while not archivo.AtEndOfStream  
    xlinea=Split(archivo.ReadLine, ";")
	if xlinea(0)=campo then
		txttxt=xlinea(1)
		exit do
	end if
	if xlinea(0)="fin" then
		exit do
	end if
loop 
archivo.Close 
Set archivo = Nothing 
Set fs = Nothing
end function

function grabalog(campo)
	sql="insert into log (motivo,idpass) values ('" & campo & "','" & session("idpass") & "')"
	set rs = oConn.Execute(sql)
end function

%>
