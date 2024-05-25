<%
if not request.cookies("LoggedIn")="TRUE" then

	Response.Redirect("NewUser.asp?Message=notloggedin")

end if
%>