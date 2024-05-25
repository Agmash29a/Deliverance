<%
'(c) 2000 for Deliverance by Riley T. Perry
'
'                                 ______________
'                           ,===:'.,            `-._
'                                `:.`---.__         `-._
'                                  `:.     `--.         `.
'                                    \.        `.         `.
'                            (,,(,    \.         `.   ____,-`.,
'                         (,'     `/   \.   ,--.___`.'
'                     ,  ,'  ,--.  `,   \.;'         `
'                      `{D, {    \  :    \;
'                        V,,'    /  /    //
'                        j;;    /  ,' ,-//.    ,---.      ,
'                        \;'   /  ,' /  _  \  /  _  \   ,'/
'                              \   `'  / \  `'  / \  `.' /
'                               `.___,'   `.__,'   `.__,'  
'
'						      **** Last Third of Main ****
%>
			</TD>
			<TD width=10>
					
				<TABLE border=0 cellPadding=0 cellSpacing=0 width="10">
					<TR>
						<TD></TD>
					</TR>
				</TABLE>
	    	
			</TD>
			<TD width=200  valign="top"> 
				<%
				'**** Cart details
				%>
				<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
					<TR>
						<TD>
							<DIV align=center class=basicfont>
							
								<TABLE border=0 cellPadding=3 cellSpacing=0 width="100%">
									<TR>
										<TD><DIV align=center>
												
												<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
													<TR height=23>
														<TD class="basicfont" valign="bottom">
															<DIV align=center><B>Shopping Cart <font color='#BD0000'><i>quickview</i></font></DIV></font>
														</td>
													</tr>
												</table>
												
											</DIV>   
										</TD>
									</TR>
								</TABLE> 
				    
							</DIV>
						</TD>
					</TR>
					<TR>
						<TD>
							
							<TABLE align=center class="outerborder" border=0 cellPadding=1 cellSpacing=0 width="100%">
								<TR>
									<TD>
										
										<TABLE align=center class="innerborder" border=0 cellPadding=1 cellSpacing=0 width="100%">
											<TR>
												<TD>
												
													<TABLE bgColor=#ffffff border=0 cellPadding=4 cellSpacing=0 width="100%">
														<TR>
															<TD class="basicfont">
																<DIV align=center><B>Number of items</B>:
																	<%
																	dim QueryText
																	dim indexMAIN
																	dim stextMAIN                   
																	dim amatchMAIN
																	dim matchesMAIN
																	dim rgxpMAIN
																	dim lastmatchMAIN		'**** result of the last match
											
																	indexMAIN = 0
											
																	if not Request.QueryString("cart")="" then
						
																		'**** pattern to look for
						
																		set rgxpMAIN = new regexp
																		rgxpMAIN.Pattern = "x"
																		rgxpMAIN.Global = true
						
																		lastmatchMAIN = 1		'**** virtual x at start of string
																		indexMAIN = 1
						
																		set matches = rgxpMAIN.Execute(Request.QueryString("cart"))
						
																		for each amatchMAIN in matches
			
																			stextMAIN = mid(Request.QueryString("cart"),lastmatchMAIN,(amatchMAIN.firstindex+1)-(lastmatchMAIN))
																			
																			QueryText = QueryText & " (select price from products where id = " & stextMAIN & ") + "  									
		
																			lastmatchMAIN = amatchMAIN.firstindex+2
																			indexMAIN=indexMAIN+1
																			
																		next
																		
																		QueryText = QueryText & "0"
																		
																		set matches=nothing
	  
																		'**** grab username
																		cm.ActiveConnection=CSTR(Application("connstring"))

																		cm.CommandText = "select 0 + (" & QueryText & ") totalprice"
																		         
																		set objrs = cm.Execute
																				
																		cm.ActiveConnection=nothing
																		
																		indexMAIN=indexMAIN-1
											
																	end if
																	%>
																	<%=indexMAIN%>
																</DIV>
															</TD>
														 </TR>
														<TR class="innerborder">
															<TD height=1>
															</TD>
														</TR>
														<TR bgColor=#ffffff>
															<TD class="basicfont">
																<DIV align=center>
																	<B>Total Cost</B>: 
																	<%
																	if cint(indexMAIN) = 0 then
				                            
																		Response.write("$0")
																		
																	else
											
																		Response.Write(formatcurrency(objrs("totalprice"),2,true,true,true))
																		
																	end if
											
																	%>
																</DIV>
															</TD>
														</TR>
													</TABLE>
													
												</TD>
											</TR>
										</TABLE>
										
									</TD>
								</TR>
							</TABLE>
							
						</TD>
					</TR>
					<TR>   
						<TD height=15>
						</TD>
					</TR> 
				</TABLE>
				<%
				'**** Top 10 Movers

				cm.ActiveConnection=CSTR(Application("connstring"))

				cm.CommandText = "select products.id,products.name product,platform.name platform, price, " _
				  & "AllPlatformsTop10.Ranking from products, AllPlatformsTop10,platform " _
				  &	" where AllPlatformsTop10.PlatformID = platform.id " _
				  &	" and AllPlatformsTop10.productid = products.id ORDER by Ranking"
				         
				set objrs = cm.Execute

				cm.ActiveConnection=nothing
				%>
				<TABLE bgColor=#ffd9d9 border=0 cellPadding=1 cellSpacing=0 width=200>
					<TR bgColor=#ffd9d9>
						<TD>
					
							<TABLE bgColor=#fffffa border=0 cellPadding=1 cellSpacing=0 width=198>
								<TR bgColor=#bd0000>
									<TD>
									
										<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
											<TR bgColor=#ffffff>
												<TD>
												
													<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
														<TR>
															<TD><IMG height=19 src="../Pictures/Top10Movers.gif" width=196></TD>
														</TR>
														<TR>
															<TD>
															
																<TABLE border=0 cellPadding=3 cellSpacing=0 width="100%">
																	<TR>
																		<TD>
																		
																			<TABLE border=0 cellPadding=1 cellSpacing=0 width="100%">
																			<%
																			'**** Iterate through all records
				                        
																			while not objrs.EOF
																				%>
																				<TR>
																					<TD width="7%" valign="top"><FONT face="Arial, Helvetica, sans-serif" size=-2><a href='<%=AppendQuery("AddToCart.asp?id=" & objrs("id"))%>'><IMG border=0 src="../pictures/GreenCheckSmall2.gif" width=10></a></font></TD>
																					<TD width="93%"><FONT face="Arial, Helvetica, sans-serif" size=-2><SPAN class=basicfont><B><%=trim(objrs("Ranking"))%>.</B></SPAN>&nbsp;<A class=links href='<%=AppendQuery("Individualgame.asp?id=" & objrs("id"))%>'><B><%=trim(objrs("product"))%> (<%=trim(objrs("Platform"))%>)</B></A> <%=formatcurrency(objrs("price"),2,true,true,true)%></font></TD>
																				</TR>
																				<%
																				objrs.movenext

																			wend
				                        
																			objrs.close
																			%>
																			</TABLE>
																		</TD>
																	</TR>
																</TABLE>
															</TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>

				<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
					<TR>
						<TD height=15>
						</TD>
					</TR>
				</TABLE>
				<%
				'**** Top 10 preorders

				cm.ActiveConnection=CSTR(Application("connstring"))

				cm.CommandText = "select products.id,products.name product,platform.name platform, price," _
				  & "AllPlatformsTop10Preorders.Ranking from products, AllPlatformsTop10Preorders,platform " _
				  &	" where AllPlatformsTop10Preorders.PlatformID = platform.id " _
				  &	" and AllPlatformsTop10Preorders.productid = products.id ORDER by Ranking"
				         
				set objrs = cm.Execute

				cm.ActiveConnection=nothing
				%>
				<TABLE bgColor=#ffd9d9 border=0 cellPadding=1 cellSpacing=0 width=200>
					<TR bgColor=#ffd9d9>
						<TD>
					
							<TABLE bgColor=#fffffa border=0 cellPadding=1 cellSpacing=0 width=198>
								<TR bgColor=#bd0000>
									<TD>
									
										<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
											<TR bgColor=#ffffff>
												<TD>
												
													<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
														<TR>
															<TD><IMG height=19 src="../Pictures/Top10Preorders.gif" width=196></TD>
														</TR>
														<TR>
															<TD>
															
																<TABLE border=0 cellPadding=3 cellSpacing=0 width="100%">
																	<TR>
																		<TD>
																		
																			<TABLE border=0 cellPadding=1 cellSpacing=0 width="100%">
																			<%
																			'**** Iterate through all records
				                        
																			while not objrs.EOF
																				%>
																				<TR>
																					<TD width="7%" valign="top"><FONT face="Arial, Helvetica, sans-serif" size=-2><a href='<%=AppendQuery("AddToCart.asp?id=" & objrs("id"))%>'><IMG border=0 src="../pictures/GreenCheckSmall2.gif" width=10></a></font></TD>
																					<TD width="93%"><FONT face="Arial, Helvetica, sans-serif" size=-2><SPAN class=basicfont><B><%=trim(objrs("Ranking"))%>.</B></SPAN>&nbsp;<A class=links href='<%=AppendQuery("Individualgame.asp?id=" & objrs("id"))%>'><B><%=trim(objrs("product"))%> (<%=trim(objrs("Platform"))%>)</B></A> <%=formatcurrency(objrs("price"),2,true,true,true)%></font></TD>
																				</TR>
																				<%
																				objrs.movenext

																			wend
				                        
																			objrs.close
																			%>
																			</TABLE>
																		</TD>
																	</TR>
																</TABLE>
															</TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>

				<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
					<TR>
						<TD height=15>
						</TD>
					</TR>
				</TABLE>
				<%
				'**** New in this month

				cm.ActiveConnection=CSTR(Application("connstring"))

				cm.CommandText = "select products.id,products.name product,platform.name platform,price" _
					 & " from products, NewInThisMonth,platform" _
					 & " where NewInThisMonth.PlatformID = platform.id" _ 
					 & " and NewInThisMonth.productid = products.id" 
				         
				set objrs = cm.Execute

				cm.ActiveConnection=nothing
				%>
				<TABLE bgColor=#ffd9d9 border=0 cellPadding=1 cellSpacing=0 width=200>
					<TR bgColor=#ffd9d9>
						<TD>
					
							<TABLE bgColor=#fffffa border=0 cellPadding=1 cellSpacing=0 width=198>
								<TR bgColor=#bd0000>
									<TD>
									
										<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
											<TR bgColor=#ffffff>
												<TD>
												
													<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
														<TR>
															<TD><IMG height=19 src="../Pictures/NewArrivals.gif" width=196></TD>
														</TR>
														<TR>
															<TD>
															
																<TABLE border=0 cellPadding=3 cellSpacing=0 width="100%">
																	<TR>
																		<TD>
																		
																			<TABLE border=0 cellPadding=1 cellSpacing=0 width="100%">
																			<%
																			'**** Iterate through all records
				                        
																			while not objrs.EOF
																				%>
																				<TR>
																					<TD width="7%" valign="top"><FONT face="Arial, Helvetica, sans-serif" size=-2><a href='<%=AppendQuery("AddToCart.asp?id=" & objrs("id"))%>'><IMG border=0 src="../pictures/GreenCheckSmall2.gif" width=10></a></font></TD>
																					<TD width="93%"><FONT face="Arial, Helvetica, sans-serif" size=-2><A class=links href='<%=AppendQuery("Individualgame.asp?id=" & objrs("id"))%>'><B><%=trim(objrs("product"))%> (<%=trim(objrs("Platform"))%>)</B></A> <%=formatcurrency(objrs("price"),2,true,true,true)%></font></TD>
																				</TR>
																				<%
																				objrs.movenext

																			wend
				                        
																			objrs.close
																			%>
																			</TABLE>
																		</TD>
																	</TR>
																</TABLE>
															</TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
				</TABLE>

				<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
					<TR>
						<TD height=15>
						</TD>
					</TR>
				</TABLE>
				<%
				'**** right column ads
				%>
				<!-- #include file="../Includes/RightColumnAdsHome.asp"-->
				<%
				'**** End main table
				%>
			</TD>
	    </TR>
	</TABLE>
	<%
	'**** Bottom Bar
	%>
	<BR><BR>
	<!-- #include file="../Includes/bottomSection.asp"-->
	
	<table width =100%>
		<tr>
			<td>
				<hr size=1>
				<center><A class=smalllinks href="<%=AppendQuery("CorporateInformation.asp")%>">Corporate information</A>&nbsp;&nbsp;|&nbsp;&nbsp;<A class=smalllinks href="<%=AppendQuery("AdvertiseOnThisSite.asp")%>">Advertise on this site</A></center>
			</td>
		</tr>
	</table>
	
</body>
<%
SET cm = NOTHING
SET objrs = NOTHING
%>
