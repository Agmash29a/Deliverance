<!--#include file="../functions/convert.asp"-->
<center class="headingfont">Search Results</center>
<table><tr height=2><td></td></tr></table>

<%
DIM Index
DIM Counter
Dim QueryTextCommand

'**** grab index value

IF request("index")="" THEN

	index=0

ELSE

	index=CINT(request("index"))
	
END IF

Counter=0

QueryTextCommand = ""

'**** Set up search text

if not request("platform") = "0" then
 
	QueryTextCommand = QueryTextCommand & " and products.platformid = " & request("platform") 

end if

if not request("genre") = "0" then

	QueryTextCommand = QueryTextCommand & " and products.genreid = " & request("genre")

end if
	
if request("Includedesc") = "on" then
	
	QueryTextCommand = QueryTextCommand & " and (products.name like '%" &  convert(Request("SearchBox")) & "%' or products.onelinedesc like '%" &  convert(Request("SearchBox")) & "%' or products.productdesc like '%" &  convert(Request("SearchBox")) & "%') "

else

	QueryTextCommand = QueryTextCommand & " and products.name like '%" &  convert(Request("SearchBox")) & "%' "
	
end if
	 
'**** grab all matches

cm.ActiveConnection=Cstr(Application("connstring"))

cm.CommandText = "Select products.id,platform.name platform,products.name,price,onelinedesc from products,platform where products.platformid=platform.id and products.deleted=0 " _
			&    QueryTextCommand
								         
set objrs = cm.Execute

cm.ActiveConnection=nothing

if objrs.eof THEN
	%>
	<center class="textfont">
	No Records Found
	</center>
	<%
else
	'**** iterate through results, sart at counter values
	%>
	<table width="100%" border="0" cellspacing="0" cellpadding="1"  height=1  class="outerborder">
	<tr><td>
	<table width="100%" border="0" cellspacing="1" cellpadding="4"  height=1  class="innerborder">		
	<tr class="headingback">
			<td width=84% align=center><b>TITLE</b></td>
			<td align=center  width=15%><b>PRICE</b></td>
			</tr>
	<%

	while NOT objrs.EOF

		IF Counter < index THEN
		
			objrs.movenext
			Counter = Counter + 1
		
		ELSE

			IF Counter < (index+10) THEN
		
				%>
			
				<TR  bgcolor="#ffffff" class="basicfont">
				<td>
				
				<table  bgcolor="#ffffff" class="basicfont" height=0 cellspacing=0>
				<tr  class="basicfont">
				<TD valign="bottom" width=12 align=left><a href='<%=AppendQuery("AddToCart.asp?id=" & objrs("id"))%>'>
				<IMG height=10 border=0 src="../pictures/GreenCheckSmall.gif" width=10 ></a>
				</TD>
				
				<TD align=left  bgcolor="#ffffff"><A class=links 
				href='<%=AppendQuery("IndividualGame.asp?id=" & objrs("id"))%>'><b><%=trim(objrs("name"))%> (<%=trim(objrs("Platform"))%>)</b></A>
				</TD>
				
				</TR>
				<%
				if not trim(objrs("onelinedesc"))="" then
					%> 
					<tr  class="basicfont">
					<td></td>
					<td align=left>
					<%=objrs("onelinedesc")%>
					</td>
					</tr>
				
					<%
				end if
				%>
				</table>
				</td>
				
				<td>
				 <%=formatcurrency(objrs("price"),2,true,true,true)%>
				</td>
				
				</tr>
				
				<%
				objrs.movenext

				Counter = Counter + 1
			
			ELSE
			
				objrs.movenext
				Counter = Counter + 1
			end if
		
		END IF
		
	wend

	%>
	</table>
	</td></tr>
	</table>
	<%

	IF Not index=0 THEN
		%>
		<a  class="speciallinks" href='<%=AppendQuery("supersearchgames.asp?index=" & index-10 & "&searchBox=" & Request("SearchBox") & "&genre=" & request("genre") & "&platform=" & request("platform") & "&Includedesc=" & request("Includedesc"))%>'>&#139;&#139;&nbsp;back</a>&nbsp;&nbsp;
		<%
	END IF

	IF (Counter-(index+10))>0 THEN
		%>
		<a class="speciallinks" href='<%=AppendQuery("Supersearchgames.asp?index=" & index+10 & "&searchBox=" & Request("SearchBox") & "&genre=" & request("genre") & "&platform=" & request("platform") & "&Includedesc=" & request("Includedesc"))%>'>next&nbsp;&#155;&#155;</a>
		<%
	END IF

end if
%>
<BR>
<hr size="1" align="center" width="80%" >
<center><a href='<%=AppendQuery("games.asp")%>' class="links" >Home</a></center>
