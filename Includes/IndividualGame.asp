<!-- #include file="../functions/convert.asp"-->
<%
cm.ActiveConnection=cstr(Application("connstring"))
 
cm.CommandText = "select datetype.name datetype,products.onelinedesc,products.SalePercentage, products.RecPrice,products.name name,distributor.name distributorname,platform.name platform, " _
& "products.Price Price, products.Recprice rrp, products.availabledatetype, products.availabledate availabledate, " _
& "products.productrating productrating, products.productdesc, genre.name genre, status.name status, " _ 
& "products.BoxshotURL from products,distributor,platform,genre,datetype,status where  " _
& "products.distributorid=distributor.id and products.platformid = platform.id and  " _
& "products.genreid = genre.id and products.availabledatetype = datetype.id and products.statusid = status.id and products.id = " & CINT(Request("id")) 
         
set objrs = cm.Execute

cm.ActiveConnection=nothing

if objrs.EOF then

	Response.Redirect(AppendQuery("games.asp"))

end if
%>
<center class="headingfont"><a href='<%=AppendQuery("AddToCart.asp?id=" & request("id"))%>'><IMG height=10 border=0 src="../pictures/GreenCheckSmall.gif" width=10 ></a>&nbsp;<%=trim(objrs("Name"))%> (<%=trim(objrs("Platform"))%>)</center>
<table><tr height=2><td></td></tr></table>
<font class="textfont">
<%
if NOT trim(objrs("BoxShotURL"))="" THEN
	%>
	<div align="center">
	<img border=0 bordercolor="#000000" src='<%=trim(objrs("BoxShotURL"))%>'></div><BR>
	<%
end if
      
'**** Check for onelinedesc or productdesc
      
      if not trim(objrs("productdesc"))="" then
		%>
		<%=trim(objrs("productdesc"))%><BR><BR>			
		<%
      elseif not trim(objrs("onelinedesc"))="" then
		%>
		<%=trim(objrs("onelinedesc"))%><BR><BR>		
		<%
      end if
      %>
		<table width="100%" border="0" cellspacing="0" cellpadding="1"  height=1  class="outerborder">
		<tr><td>

		<table width="100%" border="0" cellspacing="1" cellpadding="4"  height=1  class="innerborder">
		
		
		
		<tr>
		<td class="headingback" width="90">
		&nbsp;<b>PRICE</b>
		</td>
		<td bgcolor="#ffffff">
		<b><%=formatcurrency(objrs("price"),2,true,true,true)%></b>&nbsp;
		<%
		'*------------------*
		'* Check for sale % *
		'*------------------*
		
		if objrs("SalePercentage") <> 0 then
			%>
			<b><font color="#BD0000">On Sale!</b> was $<%=objrs("Recprice")%>, now reduced by <%=objrs("SalePercentage")%>%</font> 
			<%
		end if
		%>
		</td>
		</tr>

		<tr>
		<td class="headingback" >
		&nbsp;<b>RELEASE DATE</b>
		</td>
		<td bgcolor="#ffffff">
		<%
		
		'**** Check date type
		
		if not objrs("availabledatetype")=1 then
			%>
			<%=objrs("datetype")%>
			<%		
		else
			%>
			<%
			if objrs("availabledate") < Date then
				%>
				Now
				<%
			else
				%>
				<%=dateconverter(objrs("availabledate"))%>
				<%
			end if
			%>
			
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
		
		<tr>
		<td class="headingback" >
		&nbsp;<b>STATUS</b>
		</td>
		<td bgcolor="#ffffff">
		<%=trim(objrs("status"))%>
		</td>
		</tr>
		</td></tr></table>
		</td></tr></table><BR>
		
		<center><a href='<%=AppendQuery("AddToCart.asp?id=" & request("id"))%>'><IMG  border=0 src="../pictures/AddtoCart.gif"  ></a></center><BR></font>			
		<hr size="1" align="center" width="80%" >
		<center><a href='<%=AppendQuery("games.asp")%>' class="links" >Home</a></center>