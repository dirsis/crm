<%
	session("mail")=""
	Response.Cookies("mail")=""
	Response.Cookies("idpass")=""
	session("idpass")=""
	session.Abandon()
	response.Redirect("login.asp")
%>
