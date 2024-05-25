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
'						      **** First Third of Main ****

dim cm				'Command Object
dim objrs			'Recordset
SET cm = Server.CreateObject("ADODB.Command")

'**** Include Javascript functions and head code
%>
<link rel="stylesheet" href="../includes/Main.css" type="text/css">
<SCRIPT language=JavaScript>
<!--
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v3.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}
//-->
</SCRIPT>
</HEAD>
<%
Response.Expires=-10000
%>
<BODY bgColor=#ffffff onload="MM_preloadImages('../Pictures/PCHighlighted.gif','../Pictures/Ps1ps2Highlighted.gif','../Pictures/NintendoHighlighted.gif','../Pictures/ApplicationsHighlighted.gif','../Pictures/XboxHighlighted.gif','../Pictures/HomeHighlighted.gif','../Pictures/ElectronicsHighlighted.gif','../Pictures/DVDHighlighted.gif','../Pictures/OnSaleHighlighted.gif')" topMargin=0 class="basicfont">
	<%
	'**** Include Title Section 
	%>
	
	<TABLE border=0 cellPadding=0 cellSpacing=0 width=100%>
		<TR>
			<TD width=156>
				<img height=65 src="../Pictures/DeliveranceLogo.gif" alt="Deliverance Software - cool games">
			</TD>
			<TD>
				<%
				'**** Middle section
				%>
				<div align=center>
				
					<TABLE border=0 cellPadding=0 cellSpacing=0 height=65 width="447">
						<TR>
							<TD>
								<DIV align=center class="textfont"><b>PC Games, PlayStation Games, PS2 Games,  Xbox Games, GameBoy Games, DVDs, Electronics, News, Reviews, and Much More!</b><BR><font size=-1>...and it's easy - just click <img src="../pictures/greenchecksmall.gif"> to add to your cart.</font></DIV>
							</TD>
						</TR>
						<TR>
							<TD>&nbsp;</TD>
						</TR>
					</TABLE>
					
				</div>
			</TD>
			<TD width=157>
				
				<TABLE border=0 cellPadding=0 cellSpacing=0 height=65>
					<TR>
						<TD vAlign=bottom width=77></TD>
						<TD vAlign=bottom width=2></TD>
						<TD vAlign=bottom width=77></TD>
						<TD width=1></TD>
					</TR>
					<TR>
						<%
						'**** Include no cookie cart functions 
						%>
						<!-- #include file="../functions/AppendQuery.asp"-->
						<%
						'**** Help and View Cart 
						%>
						<TD vAlign=bottom width=77>
							<A href="<%=AppendQuery("Help.asp")%>"><IMG border=0 src="../pictures/help.gif"></A>
						</TD>
						<TD vAlign=bottom width=2>
						</TD>
						<TD vAlign=bottom width=77><A href="<%=AppendQuery("ViewCart.asp")%>">
							<IMG border=0 src="../pictures/cart.gif"></A>
						</TD>
						<TD vAlign=bottom width=1>
						</TD>
					</TR>
				</TABLE>
				
			</TD>
		</TR>
	</TABLE>


	<%
	'**** Include Navigation Bar
	%>

	<TABLE border=0 cellPadding=0 cellSpacing=0 width=100%>
		<TR>
			<td width = 50%>
			
				<TABLE border=0 cellPadding=0 cellSpacing=0 width=100%>
					<tr height=1>
						<td bgcolor=#FFD9D9>
						</td>
					</tr>
					<tr height=19>
						<td bgcolor=#BD0000>
						</td>
					
					</tr>
					<tr height=1>
						<td bgcolor=#FFD9D9>
						</td>
					</tr>
				</table>
				
			</td>
			<TD width=10>
			
				<TABLE border=0 cellPadding=0 cellSpacing=0 width=760>
					<TR>
						<TD width=50%>
							<TABLE border=0 cellPadding=0 cellSpacing=0 width=100%>
								<tr height=1>
									<td bgcolor=#FFD9D9>
									</td>
								</tr>
								<tr height=19>
									<td bgcolor=#BD0000>
									</td>
								
								</tr>
								<tr height=1>
									<td bgcolor=#FFD9D9>
									</td>
								</tr>
							</table>
							
						</td>
						<td>
							<TABLE align=center border=0 cellPadding=0 cellSpacing=0>
								<TR>
									<TD><IMG height=21 src="../Pictures/FillBar.gif" ></TD>
									<TD><A href="../<%=AppendQuery("All_Games/Games.asp")%>"  onmouseout=MM_swapImgRestore() onmouseover="MM_swapImage('Image109','','../Pictures/HomeHighlighted.gif',1)"><IMG border=0 height=21 name=Image109 src="../Pictures/Home.gif"  alt="All Games"></A></TD>
									<TD><IMG height=21 src="../Pictures/MiddleBar.gif" width=23></TD>
									<TD><A href="../<%=AppendQuery("PC_Games/Games.asp")%>" onmouseout=MM_swapImgRestore() onmouseover="MM_swapImage('Image10311','','../Pictures/PCHighlighted.gif',1)"><IMG border=0 height=21 name=Image10311 src="../Pictures/PC.gif"  alt="PC Games"></A></TD>
									<TD><IMG height=21 src="../Pictures/MiddleBar.gif" width=23></TD>
									<TD><A href="../<%=AppendQuery("PlayStation_Games/Games.asp")%>" onmouseout=MM_swapImgRestore() onmouseover="MM_swapImage('Image1144','','../Pictures/Ps1Ps2Highlighted.gif',1)"><IMG border=0 height=21 name=Image1144 src="../Pictures/ps1ps2.gif" alt="PlayStation Games" ></A></TD>
									<TD><IMG height=21 src="../Pictures/MiddleBar.gif" width=23></TD>
									<TD><A href="../<%=AppendQuery("Nintendo_Games/Games.asp")%>" onmouseout=MM_swapImgRestore() onmouseover="MM_swapImage('Image1123','','../Pictures/NintendoHighlighted.gif',1)"><IMG border=0 height=21 name=Image1123 src="../Pictures/Nintendo.gif" alt="Nintendo Games" ></A></TD>
									<TD><IMG height=21 src="../Pictures/MiddleBar.gif" width=23></TD>
									<TD><A href="../<%=AppendQuery("Xbox_Games/Games.asp")%>" onmouseout=MM_swapImgRestore() onmouseover="MM_swapImage('Image10911','','../Pictures/XboxHighlighted.gif',1)"><IMG border=0 height=21 name=Image10911 src="../Pictures/XboxHighlighted.gif" alt="Xbox Games"></A></TD>
									<TD><IMG height=21 src="../Pictures/MiddleBar.gif" width=23></TD>
									<TD><A href="../<%=AppendQuery("DVD/Games.asp")%>" onmouseout=MM_swapImgRestore() onmouseover="MM_swapImage('Image112','','../Pictures/DVDHighlighted.gif',1)"><IMG border=0 height=21 name=Image112 src="../Pictures/DVD.gif"alt="DVDs"></A></TD>
									<TD><IMG height=21 src="../Pictures/MiddleBar.gif" width=23></TD>
									<TD><A href="../<%=AppendQuery("Electronics/Games.asp")%>" onmouseout=MM_swapImgRestore() onmouseover="MM_swapImage('Image10711','','../Pictures/ElectronicsHighlighted.gif',1)"><IMG border=0 height=21 name=Image10711 src="../Pictures/Electronics.gif" alt="Electronics"></A></TD>
									<TD><IMG height=21 src="../Pictures/MiddleBar.gif" width=23></TD>
									<TD><A href="../<%=AppendQuery("Applications/Games.asp")%>" onmouseout=MM_swapImgRestore() onmouseover="MM_swapImage('Image10611','','../Pictures/ApplicationsHighlighted.gif',1)"><IMG border=0 height=21 name=Image10611 src="../Pictures/Applications.gif" alt="Applications"></A></TD>
									<TD><IMG height=21 src="../Pictures/MiddleBar.gif" width=23></TD>
									<TD><A href="../<%=AppendQuery("On_Sale/Games.asp")%>" onmouseout=MM_swapImgRestore() onmouseover="MM_swapImage('Image111','','../Pictures/OnSaleHighlighted.gif',1)"><IMG border=0 height=21 name=Image111 src="../Pictures/OnSale.gif" alt="On Sale"></A></TD>
									<TD><IMG height=21 src="../Pictures/FillBar.gif" ></TD>
	                			</TR>
	                		</TABLE>
	                					
						</TD>
						<TD width=50%>
							<TABLE border=0 cellPadding=0 cellSpacing=0 width=100%>
								<tr height=1>
									<td bgcolor=#FFD9D9>
									</td>
								</tr>
								<tr height=19>
									<td bgcolor=#BD0000>
									</td>
								
								</tr>
								<tr height=1>
									<td bgcolor=#FFD9D9>
									</td>
								</tr>
							</table>
							
						</td>
					</TR>
				</TABLE>
	                		
			</TD>
			<td width=50%>
				
				<TABLE border=0 cellPadding=0 cellSpacing=0 width=100%>
					<tr height=1>
						<td bgcolor=#FFD9D9>
						</td>
					</tr>
					<tr height=19>
						<td bgcolor=#BD0000>
						</td>
					
					</tr>
					<tr height=1>
						<td bgcolor=#FFD9D9>
						</td>
					</tr>
				</table>
				
			</td>
		</TR>  
	</TABLE>

	<%
	'**** Insert space between Header and main table
	%>
	
	<TABLE border=0 cellPadding=0 cellSpacing=0 height=15 width=100%>
		<TR>
			<TD>
			</TD>
		</TR>
	</TABLE>
	
	<%
	'*------------*
	'* Main Table *
	'*------------*
	%>
	
	<TABLE border=0 cellPadding=0 cellSpacing=0 width=100%>
		<TR> 
			<TD vAlign=top width=130>
				
				<TABLE align=center class="outerborder" border=0 cellPadding=1 cellSpacing=0>
					<TR>
						<TD>
							
							<TABLE align=center class="innerborder" border=0 cellPadding=1 cellSpacing=0 width="100%">
								<TR>
									<TD>
										
										<TABLE bgColor=#ffffff border=0 cellPadding=0 cellSpacing=0 width="99%">
											<TR bgColor=#ffffff>
												<TD>
													
													<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
														<TR>
															<TD><form name="KeySearch" action="<%=AppendQuery("Search.asp")%>" method="post">
																
																<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
																	<TR>
																		<TD><IMG height=18 src="../Pictures/SearchBar.gif" width=88></TD>
																		<TD><input type="image" name="button" value="Submit" src="../Pictures/Go.gif"></TD>
																	</TR>
																</TABLE>
															</TD>
														</TR>
													</TABLE>
													
													<TABLE border=0 cellPadding=3 cellSpacing=0 width="100%">
														<TR class="innerborder" >
															<TD height=1></TD>
														</TR>
														<TR >
															<TD><DIV align=center><INPUT name=SearchBox class="textbox"  maxlength=255></DIV></TD>
														</TR></form>
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
						<TD height=15></TD>
					</TR>
				</TABLE>
				
				<TABLE class="outerborder" border=0 cellPadding=1 cellSpacing=0 width="100%">
					<TR class="outerborder">
						<TD>
						
							<TABLE class="innerborder" border=0 cellPadding=1 cellSpacing=0 width="100%">
								<TR class="innerborder">
									<TD>
									
										<TABLE  border=0 cellPadding=0 cellSpacing=0 width="100%">
											<TR bgColor=#ffffff>
												<TD>
												
													<TABLE bgColor=#ffffff border=0 cellPadding=0 cellSpacing=0 width="100%">
														<TR>
															<TD>
																
																<TABLE border=0 cellPadding=4 cellSpacing=0 width="100%">
																	<TR>
																		<TD valign = middle>
																			
																			<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
																				<TR>
																					<TD><FONT face="Arial, Helvetica, sans-serif" size=-2><DIV align=left><B class=basicfont>&nbsp;MEMBERS</B></DIV></font>
																					</TD>
																				</TR>
																				<TR class="innerborder">
																					<TD height=1></TD>
																				</TR>
																				<TR>
																					<TD><FONT face="Arial, Helvetica, sans-serif" size=-2>&#8226;&nbsp;<A href="<%=AppendQuery("YourMembership.asp")%>"><SPAN class=links>Member FAQ</SPAN></A><BR>&#8226;&nbsp;<A class=links href="<%=AppendQuery("LoginEdit.asp")%>"><SPAN class=links>Edit / View Details</SPAN></A> <BR>&#8226;&nbsp;<A class=links href="<%=AppendQuery("NewUser.asp")%>"><SPAN class=links><B><I>New</I></B>Membership</SPAN></A> 
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
							
						</TD>
					</TR>
				</TABLE>
				
				<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
					<TR>
						<TD height=15></TD>
					</TR>
				</TABLE>
				
				<TABLE class="outerborder" border=0 cellPadding=1 cellSpacing=0 width="100%">
					<TR class="outerborder">
						<TD>
						
							<TABLE class="innerborder" border=0 cellPadding=1 cellSpacing=0 width="100%">
								<TR class="innerborder">
									<TD>
									
										<TABLE bgColor=#fffff0 border=0 cellPadding=0 cellSpacing=0 width="100%">
											<TR bgColor=#ffffff>
												<TD>
												
													<TABLE bgColor=#ffffff border=0 cellPadding=0 cellSpacing=0 width="100%">
														<TR>
															<TD>
																
																<TABLE border=0 cellPadding=4 cellSpacing=0 width="100%">
																	<TR>
																		<TD>
																			
																			<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
																				<TR>
																					<TD><FONT face="Arial, Helvetica, sans-serif" size=-2><DIV align=left><B class=basicfont>&nbsp;ORDERS</B></DIV></font>
																					</TD>
																				</TR>
																				<TR class="innerborder">
																					<TD height=1></TD>
																				</TR>
																				<TR>
																					<TD><FONT face="Arial, Helvetica, sans-serif" size=-2>&#8226;&nbsp;<A href="<%=AppendQuery("OrderFAQ.asp")%>"><SPAN class=links>Order FAQ</SPAN></A><BR>&#8226;&nbsp;<A href="<%=AppendQuery("LoginTrackOrder.asp")%>"><SPAN class=links>Track an Order</a></SPAN><BR>&#8226;&nbsp;<A href="<%=AppendQuery("DeliveryInfo.asp")%>"><SPAN class=links>Delivery Info.</SPAN></A></font></TD>
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
							
						</TD>
					</TR>
				</TABLE>
				
				<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
					<TR>
						<TD height=15></TD>
					</TR>
				</TABLE>
				
				<TABLE class="outerborder" border=0 cellPadding=1 cellSpacing=0 width="100%">
					<TR class="outerborder">
						<TD>
						
							<TABLE class="innerborder" border=0 cellPadding=1 cellSpacing=0 width="100%">
								<TR class="innerborder">
									<TD>
									
										<TABLE bgColor=#fffff0 border=0 cellPadding=0 cellSpacing=0 width="100%">
											<TR bgColor=#ffffff>
												<TD>
												
													<TABLE bgColor=#ffffff border=0 cellPadding=0 cellSpacing=0 width="100%">
														<TR>
															<TD>
																
																<TABLE border=0 cellPadding=4 cellSpacing=0 width="100%">
																	<TR>
																		<TD>
																			
																			<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
																				<TR>
																					<TD><FONT face="Arial, Helvetica, sans-serif" size=-2><DIV align=left><B class=basicfont>&nbsp;COMMUNITY</B></DIV></font>
																					</TD>
																				</TR>
																				<TR class="innerborder">
																					<TD height=1></TD>
																				</TR>
																				<TR>
																					<TD><FONT face="Arial, Helvetica, sans-serif" size=-2>&#8226;&nbsp;<A href="<%=AppendQuery("links.asp")%>"><SPAN class=links>Links</SPAN></a></font>
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
						</TD>
					</TR>
				</TABLE>
				
				<TABLE border=0 cellPadding=0 cellSpacing=0 width="100%">
					<TR>
						<TD height=15></TD>
					</TR>
				</TABLE>
				  
				<!-- #include file="../Includes/LeftColumnAdsHome.asp"-->
      
			</TD>
			<TD width=10>
				
				<TABLE border=0 cellPadding=0 cellSpacing=0 width="10">
					<TR>
						<TD></TD>
					</TR>
				</TABLE>
    
			</TD>
			<TD vAlign=top>
				<%
				'**** width minimum table
				%>
				<TABLE border=0 cellPadding=0 cellSpacing=0 width="410">
					<TR>
						<TD></TD>
					</TR>
				</TABLE>
