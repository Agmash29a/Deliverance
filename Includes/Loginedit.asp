<!--#include file="../Includes/Constants.asp"-->
<%
dim responsetext

select case request("message")
	
	case "badusernameorpassword"

		responsetext = "Bad User name or Password, try again or create a new user<BR>"
	
end select
%>

<center class="headingfont">Log in to our <i>Secure</i> Server</center>
<table><tr height=2><td></td></tr></table>
<center class="textfont"><a class="invalidtext"><%=responsetext%></a>Please enter your details below</center>
<table><tr height=2><td></td></tr></table>

<div align=center><table width="410" border="0" cellspacing="0" cellpadding="1"  height=1  class="outerborder">
<form action="<%=AppendQuery(SSLURL & "loginedit.asp")%>" method="post">
<tr><td>
<table width="100%" border="0" cellspacing="1" cellpadding="2"  height=1  class="innerborder">
<tr>
<td class="headingback" width="74">
&nbsp;<b>USERNAME</b>
</td>
<td bgcolor="#FCF5AC" align="center">
<input type ="textbox" name="username" class="textbox" value=''  maxlength=50>
</td>
</tr>

<tr>
<td class="headingback">
&nbsp;<b>PASSWORD</b>
</td>
<td bgcolor="#FCF5AC" align="center">
<input name="password" type="password" class="textbox"  maxlength=50>
</td>
</tr>

</table></td></tr></table></div><BR>
<center><input type="submit" name="Login" value="Login"></form></center>
<center><b><a href="<%=AppendQuery("NewUser.asp")%>" class="textfontlinks">Click here if you are not a member</a></b></center><BR>
<hr size="1" align="center" width="80%" >
<center><a href='<%=AppendQuery("games.asp")%>' class="links" >Home</a></center><BR>