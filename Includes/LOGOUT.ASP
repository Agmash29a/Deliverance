<%
Response.cookies("LoggedIn")="FALSE"
session("cartcounter")=""
Response.Redirect(AppendQuery("goodbye.asp"))
%>
