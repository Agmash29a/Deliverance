<%
if request.cookies("LoggedIn")="TRUE" then

	cm.ActiveConnection=Cstr(Application("connstring"))

	cm.CommandText = "select orderitems.id from orderitems,orders where orderid = " &  request("invoicenumber") _	
	& "and orders.userid = " & Request.Cookies("userid") & " and orderitems.id = " & request("id") & " and orderitems.orderid = orders.id"

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