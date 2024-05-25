<!--#include file="../functions/Convert.asp"-->
<!--#include file="../Includes/Constants.asp"-->
<%
dim responsetext

select case request("message")

	case "notloggedin"
	
		responsetext = "You must become a member to use that feature<BR>"
	
	case "badusernameorpassword"

		responsetext = "Bad User name or Password, try again or create a new user<BR>"
	
	case "mustlogintobuy"

		responsetext = "You must become a member to complete the buying process<BR>"
		
	case "passwordsdontmatch"

		responsetext = "The passwords did not match<BR>"
		
	case "usernametaken"

		responsetext = "That username is already taken<BR>"
	
		
end select

'**** if not password dont match, wipe sessions

if not request("message") = "passwordsdontmatch" or responsetext = "That username is already taken" then

	Session("Firstname")=""
	Session("lastname")=""
	Session("email")=""
	Session("PhoneNumber")=""
	Session("username")=""
				
end if
%>

<center class="headingfont">New User</center>
<table><tr height=2><td></td></tr></table>
<center class="textfont"><a class="invalidtext"><%=responsetext%></a>Please enter your details below</center>
<table><tr height=2><td></td></tr></table>
<div align=center>
<table width="410" border="0" cellspacing="0" cellpadding="1"  height=1  class="outerborder">
<form action="<%=AppendQuery(SSLURL & "AddNewUser.asp")%>" method="post">
<tr><td>

<table width="410" border="0" cellspacing="1" cellpadding="2"  height=1  class="innerborder">
<tr>
<td class="headingback" width="130">
&nbsp;<b>FIRST NAME</b>
</td>
<td bgcolor="#FCF5AC"  align="center">
<input type ="textbox" name="Firstname" class="textbox" value='<%=valueconvert(session("Firstname"))%>' maxlength=50>
</td>
</tr>
<tr>
<td class="headingback">
&nbsp;<b>LAST NAME</b>
</td>
<td bgcolor="#FCF5AC"  align="center">
<input type ="textbox" name="lastname" class="textbox" value='<%=valueconvert(session("lastname"))%>' maxlength=50>
</td>
</tr>

<tr>
<td class="headingback">
&nbsp;<b>EMAIL</b>
</td>
<td bgcolor="#FCF5AC"  align="center">
<input type ="textbox" name="email" class="textbox" value='<%=valueconvert(session("email"))%>' maxlength=50>
</td>
</tr>


<tr>
<td class="headingback">
&nbsp;<b>PHONE NUMBER</b>
</td>
<td bgcolor="#FCF5AC"  align="center">
<input type ="textbox" name="PhoneNumber" class="textbox" value='<%=valueconvert(session("PhoneNumber"))%>'  maxlength=50>
</td>
</tr>
</table></td></tr></table><BR>


<table width="410" border="0" cellspacing="0" cellpadding="1"  height=1  class="outerborder">
<tr><td>
<table width="410" border="0" cellspacing="1" cellpadding="2"  height=1  class="innerborder">
<tr>
<td class="headingback" width="130">
&nbsp;<b>USERNAME</b>
</td>
<td bgcolor="#FCF5AC"  align="center">
<input type ="textbox" name="username" class="textbox" value='<%=valueconvert(session("username"))%>'  maxlength=50>
</td>
</tr>

<tr>
<td class="headingback">
&nbsp;<b>PASSWORD</b>
</td>
<td bgcolor="#FCF5AC"  align="center">
<input name="password" type="password" class="textbox"  maxlength=50>
</td>
</tr>

<tr>
<td class="headingback">
&nbsp;<b>CONFIRM PASSWORD</b>
</td>
<td bgcolor="#FCF5AC"  align="center">
<input name="Confirmpassword" type="password" class="textbox" maxlength=50>
</td>
</tr>

</table></td></tr></table></div><BR>
<center><input type="submit" name="Create" value="Create User"></form></center>
<center><font class="smallFont">* All user details are strictly confidential, we will not dislcose your details to any other organisation</font></center>
<BR>
<hr size="1" align="center" width="80%" >
<center><a href='<%=AppendQuery("games.asp")%>' class="links" >Home</a></center><BR>