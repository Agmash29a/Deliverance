<!-- #include file="../functions/convert.asp"-->
<%
SET cm = Server.CreateObject("ADODB.Command")


    if not Request.QueryString("cart") = "" then
		
		dim NewText	'**** updated Querystring
		dim stext   '**** product id
		dim Shippingtestsum '**** total for available orders
		
			%>
			
			<center  class="headingfont">Shopping Cart</center>
			
			<table><form name="cart" action="<%=AppendQuery("UpdateCart.asp")%>" method="post"><tr height=2><td></td></tr></table>
			<table width="100%" border="0" cellspacing="0" cellpadding="1"  height=1  class="outerborder">
			<tr><td>

			<table width="100%" border="0" cellspacing="1" cellpadding="4"  height=1  class="innerborder">
				<tr class="headingback">
				<td align=center  width="30"><b>REMOVE</b></td>
				<td  align=center><b>TITLE</b></td>
				<td align=center  width="100"><b>RELEASE DATE</b></td>
				<td align=center  width="60"><b>PRICE</b></td>
				</tr>
				<%
				dim amatch
				dim matches
				dim rgxp
				dim lastmatch		'**** result of the last match
				dim preordercount
				
				'**** pattern to look for
				
				set rgxp = new regexp
				rgxp.Pattern = "x"
				rgxp.Global = true
				
				lastmatch = 1		'**** virtual x at start of string
				index = 1
				preordercount = 1
				
				set matches = rgxp.Execute(Request.QueryString("cart"))
				
				for each amatch in matches
		
					stext = mid(Request.QueryString("cart"),lastmatch,(amatch.firstindex+1)-(lastmatch))
					
					cm.ActiveConnection=Cstr(Application("connstring"))
					cm.CommandText = "select datetype.name datetype,products.name,availabledate,availabledatetype,price,platform.name platform from products,platform,datetype where " _
							   & " platform.id = products.platformid and availabledatetype = datetype.id and products.id = (" & stext & ")"
					
					set objrs = cm.Execute
					cm.ActiveConnection=nothing
					%>

						<%
						if objrs("availabledatetype")=1 then
					
							if objrs("availabledate") < Date then
								%>
								<tr   bgcolor="#fffff0"   class="basicfont">
								<td align=center>
								<input type="checkbox" name="id<%=index%>">
								</td>
								<td>
									<%=objrs("name")%>&nbsp;(<%=objrs("platform")%>)
								</td>
								<td>
								<center>
									Now
								</center>
								</td>
								<td>
									<%=formatcurrency(objrs("price"),2,true,true,true)%>
								<%
							else
								%>
								<tr   bgcolor="#fffff0"   class="basicfont">
								<td align=center>
								<input type="checkbox" name="id<%=index%>">
								</td>
								<td>
									<%=objrs("name")%>&nbsp;(<%=objrs("platform")%>)
								</td>
								<td>
								<center>
									<%=Dateconverter(objrs("availabledate"))%>
								</center>
								</td>
								<td>
									<%=formatcurrency(objrs("price"),2,true,true,true)%>
								<%
							end if
							
							Shippingtestsum = Shippingtestsum + cdbl(objrs("price"))
							
						else
							%>
							<tr   bgcolor="#ffffff"   class="basicfont">
							<td align=center>
							<input type="checkbox" name="id<%=index%>">
							</td>
							<td>
								<%=objrs("name")%>&nbsp;(<%=objrs("platform")%>)
							</td>
							<td>
							<center>
								<%=objrs("datetype")%>
							</td>
							<td>
								<%=formatcurrency(objrs("price"),2,true,true,true)%>
							<%
							
							preordercount=preordercount+1
							
						end if
						%>
						</td>
						</tr>		
					<%

					sum = sum + cdbl(objrs("price"))
					
					lastmatch = amatch.firstindex+2
					index=index+1
					
				next
				%>
			<tr   bgcolor="#ffffff"  class="basicfont">
			<td colspan=2 class="basicfont"><center>* All prices include GST</center></td>
			<td   class="headingback" ><center>
				<b>SUB TOTAL</b></center>
			</td>
			<td  >
				<%=formatcurrency(sum,2,true,true,true)%>
			</td>
			</tr>
	
			</table>
			</td></tr></table>
			<BR>
			<center><input type="submit" name="Update" value="Delete from Cart"></center>
			</form>
			
			
			<form name="nextstep" method="post" action="<%=AppendQuery("checkout.asp")%>">
			<table width="100%" border="0" cellspacing="0" cellpadding="1"  height=1  class="outerborder">
			<tr><td>
			<table width="100%" border="0" cellspacing="1" cellpadding="4"  height=1  class="innerborder">
			<tr class="headingback">
			<td  align=center><b>SHIPPING TYPE</b></td>
			<td align=center width="100"><b>COURIER</b></td>
			<td align=center width="60"><b>PRICE</b></td>
	
					<tr  bgcolor="#ffffff" class="basicfont">
					<td><center>
					<i>Express</i>
					</center>
					</td>
					<td><center>TNT</center>
					</td>
					<td>
					<%
					
					'**** Check if all preorders or if shipping cost > 150
					
					if preordercount = index then
						%>
						<center><i>N/A</i></center>
						<%
					else
					
						if Shippingtestsum > 150 then					
							%>
							<center><i>FREE!*</i></center>
							<%
						else
							%>
							$5.95*
							<%
						end if
					end if
					%>
					</td>
					</tr>
	
					
					<tr bgcolor="#ffffff"  class="basicfont">
					<td><center>* Only applies to available items</center>
					</td>
					<td class="headingback"><center><b>
					TOTAL</b></center>
					</td >
					<td>
					<%
					'**** If all preorders OR total > $150 then free shipping
					 
					if preordercount = index or Shippingtestsum > 150 then
					
					
						response.write(formatcurrency(sum,2,true,true,true))
						
					else
						
						response.write(formatcurrency((sum + 5.95),2,true,true,true))
						
					end if
					%>
			</td>
			</tr>
				
			</table>
			</td></tr></table><BR>
			<div class="basicfont" align=center>The above total does <b>NOT</b> include the shipping cost of preorders, remember though that 
			shipping is free on all preorders over $50</div>
			<input type="hidden" name="shippingtype" value="1">

			<BR>
				<table width=100%>
				<tr>
					<td width =33%></td>
					<td>
					</td>
				
					<td align=right width =33%><input type="image" src='../pictures/nextstep.gif' border=0></a></td>
				
				</tr>
				</table><BR>
				<hr size="1" align="center" width="80%" >
				<center><a href='<%=AppendQuery("games.asp")%>' class="links" >Home</a></center></form>
			<%
	ELSE
		%>
		<center class="headingfont">There Are No Items In Your Cart</center><center class="textfont">
		<BR>
		To add something to your cart just click the <IMG height=10 
            src="../Pictures/GreenCheckSmall.gif" width=10> icon that's next to a game title that you want.
		</center>
		<BR>
		<hr size="1" align="center" width="80%" >
		<center><a href='<%=AppendQuery("games.asp")%>' class="links" >Home</a></center>
		<%
	END IF

set objrs2=NOTHING
%>


