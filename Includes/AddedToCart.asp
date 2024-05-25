<%
cm.ActiveConnection=cstr(Application("connstring"))
 
cm.CommandText = "select datetype.name datetype,products.onelinedesc,products.name name,distributor.name distributorname,platform.name platform, " _
& "products.Price Price, products.Recprice rrp, products.availabledatetype, products.availabledate availabledate, " _
& "products.productrating productrating, products.productdesc, genre.name genre, " _ 
& "products.BoxshotURL from products,distributor,platform,genre,datetype where  " _
& "products.distributorid=distributor.id and products.platformid = platform.id and  " _
& "products.genreid = genre.id and products.availabledatetype = datetype.id and products.id = " & CINT(Request("id")) 
         
set objrs = cm.Execute

cm.ActiveConnection=nothing

if objrs.EOF then
	
	'**** not found, user error
	Response.Redirect(AppendQuery("games.asp"))

end if
%>
<center class="headingfont">Game Added to Cart</center>
<table><tr height=2><td></td></tr></table>


		<table width="100%" border="0" cellspacing="0" cellpadding="1"  height=1  class="outerborder">
		<tr><td>

		<table width="100%" border="0" cellspacing="1" cellpadding="4"  height=1  class="innerborder">
		


		<tr>
		<td class="headingback" width="76">
		&nbsp;<b>NAME</b>
		</td>
		<td bgcolor="#ffffff">
		<%=objrs("Name")%>
		</td>
		</tr>
		
				


		<tr>
		<td class="headingback" >
		&nbsp;<b>PLATFORM</b>
		</td>
		<td bgcolor="#ffffff">
		<%=objrs("Platform")%>
		</td>
		</tr>		
		<tr>
		
		
		
		<td class="headingback" >
		&nbsp;<b>PRICE</b>
		</td>
		<td bgcolor="#ffffff">
		<%=formatcurrency(objrs("price"),2,true,true,true)%>
		</td>
		</tr>
		
		<tr>
		<td class="headingback" >
		&nbsp;<b>AVAIL. DATE</b>
		</td>
		<td bgcolor="#ffffff">
		<%
		
		'**** Check date type

		if objrs("availabledatetype")=1 then
					
			if objrs("availabledate") < Date then
				%>
				Now
				<%
			else
				%>
				<%=objrs("availabledate")%>
				<%
			end if
						
		else
		
			%>
			<%=objrs("datetype")%>
			<%
					
		end if
										
		%>
		</td>
		</tr>
		
	
		<tr>
		<td class="headingback" >
		&nbsp;<b>RATING</b>
		</td>
		<td bgcolor="#ffffff">
		<%=trim(objrs("productrating"))%>
		</td>
		</tr>		


		<tr>
		<td class="headingback" >
		&nbsp;<b>GENRE</b>
		</td>
		<td bgcolor="#ffffff">
		<%=trim(objrs("genre"))%>
		</td>
		</tr>
		</td></tr></table>
		</td></tr></table><BR>
		<center><A href="<%=AppendQuery("ViewCart.asp")%>" valign=bottom><IMG border=0 src="../pictures/cart.gif"></A></center><BR>
		<center class="Textfont">Click view cart to complete buying process</center><BR>
		
	<hr size="1" align="center" width="80%" >
	<center><a href='<%=AppendQuery("games.asp")%>' class="links" >Home</a></center>			
	