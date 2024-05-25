<!--#include file="../functions/Convert.asp"-->
<%
dim responsetext

select case request("message")
		
	case "passwordsdontmatch"

		responsetext = "The passwords did not match<br>"
		
	case "usernametaken"

		responsetext = "That username is already taken<br>"
	
end select

cm.ActiveConnection=Cstr(Application("connstring"))

cm.CommandText = "select * from usr where id = " & CINT(Request.Cookies("userid"))
								         
set objrs = cm.Execute
cm.ActiveConnection=nothing

'**** if not password dont match, wipe sessions

if not request("message") = "passwordsdontmatch" or responsetext = "That username is already taken" then

	Session("firstname")=""
	Session("lastname")=""
	Session("email")=""
	Session("PhoneNumber")=""
	Session("username")=""
				
end if
%>

<center class="headingfont">Edit User Details</center>
<table><tr height=2><td></td></tr></table>
<center><a class="invalidtext"><%=responsetext%></a><font class="textfont">Please enter your details below</font></center>
<table><tr height=2><td></td></tr></table>
<table width="100%" border="0" cellspacing="0" cellpadding="1"  height=1  class="outerborder">
<form action="<%=AppendQuery("edituserupdate.asp")%>" type="post">
<tr><td>

<table width="100%" border="0" cellspacing="1" cellpadding="2"  height=1  class="innerborder">
<tr>
<td class="headingback" width=33%>
&nbsp;<b>FIRST NAME</b>
</td>
<td bgcolor="#FCF5AC" align="center">
<%
if 	Session("Firstname")="" then
	%>
	<input type ="textbox" name="Firstname" class="textbox" value='<%=valueconvert(objrs("GivenName"))%>'>
	<%
else
	%>
	<input type ="textbox" name="Firstname" class="textbox" value='<%=valueconvert(session("Firstname"))%>'>
	<%
end if
	
%>

</td>
</tr>
<tr>
<td class="headingback">
&nbsp;<b>LAST NAME</b>
</td>
<td bgcolor="#FCF5AC" align="center">
<%
if 	Session("lastname")="" then
	%>
	<input type ="textbox" name="lastname" class="textbox" value='<%=valueconvert(objrs("LastName"))%>'>
	<%
else
	%>
	<input type ="textbox" name="lastname" class="textbox" value='<%=valueconvert(session("lastname"))%>'>
	<%
end if
	
%>

</td>
</tr>

<tr>
<td class="headingback">
&nbsp;<b>EMAIL</b>
</td>
<td bgcolor="#FCF5AC" align="center">
<%
if 	Session("email")="" then
	%>
	<input type ="textbox" name="email" class="textbox" value='<%=valueconvert(objrs("email"))%>'>
	<%
else
	%>
	<input type ="textbox" name="email" class="textbox" value='<%=valueconvert(session("email"))%>'>
	<%
end if
	
%>

</td>
</tr>


<tr>
<td class="headingback">
&nbsp;<b>PHONE NUMBER</b>
</td>
<td bgcolor="#FCF5AC" align="center">
<%
if Session("PhoneNumber")="" then
	%>
	<input type ="textbox" name="PhoneNumber" class="textbox" value='<%=valueconvert(objrs("phonenumber"))%>'>
	<%
else
	%>
	<input type ="textbox" name="PhoneNumber" class="textbox" value='<%=valueconvert(session("PhoneNumber"))%>'>
	<%
end if
	
%>

</td>
</tr>
</table></td></tr></table><BR>


<table width="100%" border="0" cellspacing="0" cellpadding="1"  height=1  class="outerborder">
<tr><td>
<table width="100%" border="0" cellspacing="1" cellpadding="2"  height=1  class="innerborder">
<tr>
<td class="headingback" width="33%">
&nbsp;<b>USERNAME</b>
</td>
<td bgcolor="#FCF5AC" align="center">
<%
if 	Session("username")="" then
	%>
	<input type ="textbox" name="username" class="textbox" value='<%=valueconvert(objrs("username"))%>'>
	<%
else
	%>
	<input type ="textbox" name="username" class="textbox" value='<%=valueconvert(session("username"))%>'>
	<%
end if
	
%>

</td>
</tr>

<tr>
<td class="headingback">
&nbsp;<b>PASSWORD</b>
</td>
<td bgcolor="#FCF5AC" align="center">
<input  name="password" type="password" class="textbox">
</td>
</tr>

<tr>
<td class="headingback">
&nbsp;<b>CONFIRM PASSWORD</b>
</td>
<td bgcolor="#FCF5AC" align="center">
<input  name="Confirmpassword" type="password" class="textbox">
</td>
</tr>

</table></td></tr></table><BR>

<table width=100%>
<tr>
	<td width =33%></td>
	<td><center>
	<input type="submit" value="Submit New Details"></center>
	</td>
	<td align=right  width =33%></td>
	</tr>
</table>
</form>
<%
	Session("Firstname")=""
	Session("lastname")=""
	Session("email")=""
	Session("PhoneNumber")=""
	Session("username")=""
%>
<hr size="1" align="center" width="80%" >
<center><a href='<%=AppendQuery("games.asp")%>' class="links" >Home</a></center><BR>