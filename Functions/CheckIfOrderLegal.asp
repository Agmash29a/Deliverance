<%
if request.cookies("LoggedIn")="TRUE" then

	cm.ActiveConnection=Cstr(Application("connstring"))

	cm.CommandText = "select id from orders where id = " &  request("invoicenumber") _	
	& " and userid = " & Request.Cookies("userid")

	set objrs = cm.Execute
	cm.ActiveConnection=NOTHING
	
	'**** User cannot access this page
	
	if objrs.eof then
	
		'**** Access denied
		
		Response.Redirect("games.asp")
	
	end if
	
else

	Response.Redirect("NewUser.asp?Message=notloggedin")

end if
%>