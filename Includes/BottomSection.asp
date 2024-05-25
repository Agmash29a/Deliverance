<BR><BR>
<div align=center>
	<TABLE border=0 cellPadding=0 cellSpacing=0 width="760"> 
		<tr>
			<td width=290 valign=top>
				
				<TABLE class="outerborder" border=0 cellPadding=1 cellSpacing=0 width="290">  
					<TR class="outerborder">
						<TD valign=top>
							<form action="<%=AppendQuery("SuperSearchGames.asp")%>" method="post" id=form1 name=form1>
							
							<TABLE class="innerborder" border=0 cellPadding=1 cellSpacing=0 width="100%">
								<tr>
									<TD>
									
										<TABLE bgColor=#fffff0 border=0 cellPadding=0 cellSpacing=0 width="100%" >
											<tr>
												<td>
												
													<TABLE  class="headingback" border=0 cellPadding=0 cellSpacing=0 class=basicfont width="100%">
														<TR>
															<TD><IMG height=18 src="../Pictures/AdvancedSearchBar.gif"></TD>
															<TD align=right><input type="image" name="button" value="Submit"  src="../Pictures/Go.gif" height=18 ></TD>
														</TR>
													</TABLE>
									
												</td>
											</tr>
											<TR class="innerborder">
												<TD height=1></TD>
											<tr>
												<td>
													<table class="headingback" cellpadding=0>
														<tr>
															<td>
															</td>
														<tr class="basicfont">
															<td width=7%>&nbsp;<b>Genre</b></td>
															<td>		
																<select name="Genre" class="basicfont"  style="width: 100%">
																<option value="0">All Genres</option>
																<%
																'**** grab all Genres
																dim objrsGenre
																	
																cm.ActiveConnection=Cstr(Application("connstring"))

																cm.CommandText = "Select id,name from genre"
																								         
																set objrsGenre = cm.Execute

																cm.ActiveConnection=nothing
																
																while not objrsGenre.eof
																	%>
																	<option value="<%=objrsGenre("ID")%>"><%=objrsGenre("name")%></option>
																	<%
																	objrsGenre.movenext
																wend
																
																set objrsGenre = nothing
																%>
																</select>
															</TD>			
														</tr>
														<tr class="basicfont">
															<TD>&nbsp;Platform</td>
															<td>
																<select name="Platform" class="basicfont"  style="width: 100%">
																	<option value="0">All Platforms</option>
																	<option value="3">PC</option>
																	<option value="5">Playstation</option>
																	<option value="4">Playstation 2</option>
																	<option value="1">N64</option>
																	<option value="7">Dreamcast</option>
																	<option value="8">Mac</option>
																	<option value="11">GBC</option>
																	<option value="12">GBA</option>
																	<option value="2">Xbox</option>
																	<option value="13">DVD</option>
																	<option value="14">Electronics</option>
																	<option value="15">Game Cube</option>
																</select>
															</td>
														</tr>		 
														<tr class="basicfont">
															<TD>&nbsp;Keyword</td>
															<td><input class="basicfont"  style="width: 100%" type=text name=Searchbox></td>  
														</TR>
														<tr class="basicfont">
															<TD colspan=2>
																<center><input class="basicfont" type=checkbox name=Includedesc>&nbsp;Include product description in keyword search</center> 
															</td>  
														</TR>
														</form>
													</table>
													
												</td>
											</tr>
										</table>	
												
									</td>
								</tr>
							</TABLE> 
									       
						</TD>
					</tr>
			    </TABLE>
			           
			</td>
			<td width=10></td>
			<td width = 460>
				
				<table class=basicfont width=100%> 
					<tr>
						<td width=33% class=textfont><b>Show me all the...</b></td>
						<td width=33% class=textfont><b>Take me to the...</b></td>
						<td width=33% class=textfont><b>I want to...</b></td>
					</tr>
					<tr>
						<td valign=top>
							&#8226;&nbsp;<A href="<%=AppendQuery("SuperSearchGames.asp?Genre=3&platform=0")%>"><SPAN class=links>Strategy Games</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("SuperSearchGames.asp?Genre=2&platform=0")%>"><SPAN class=links>RPG Games</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("SuperSearchGames.asp?Genre=1&platform=0")%>"><SPAN class=links>Action Games</SPAN></a><BR>			
							&#8226;&nbsp;<A href="<%=AppendQuery("SuperSearchGames.asp?Genre=5&platform=0")%>"><SPAN class=links>Kids Games</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("SuperSearchGames.asp?Genre=4&platform=0")%>"><SPAN class=links>Puzzle Games</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("SuperSearchGames.asp?Genre=6&platform=0")%>"><SPAN class=links>Sports Games</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("SuperSearchGames.asp?Genre=7&platform=0")%>"><SPAN class=links>Management Sims</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("SuperSearchGames.asp?Genre=8&platform=0")%>"><SPAN class=links>Racing Games</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("SuperSearchGames.asp?Genre=9&platform=0")%>"><SPAN class=links>Adventure Games</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("SuperSearchGames.asp?Genre=10&platform=0")%>"><SPAN class=links>Flight Sims</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("SuperSearchGames.asp?Genre=11&platform=0")%>"><SPAN class=links>Operating Systems</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("SuperSearchGames.asp?Genre=12&platform=0")%>"><SPAN class=links>Productivity Software</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("SuperSearchGames.asp?Genre=0&platform=13")%>"><SPAN class=links>DVDs</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("SuperSearchGames.asp?Genre=0&platform=14")%>"><SPAN class=links>Electronic Stuff</SPAN></a><BR>
						</td>
						<td valign=top>
							&#8226;&nbsp;<A href="<%=AppendQuery("../All_Games/Games.asp")%>"><SPAN class=links>Deliverance Homepage</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("../PC_Games/Games.asp")%>"><SPAN class=links>PC Homepage</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("../PlayStation_Games/Games.asp")%>"><SPAN class=links>PS One / PS2 Homepage</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("../Nintendo_Games/Games.asp")%>"><SPAN class=links>Nintendo Homepage</SPAN></a><BR>			
							&#8226;&nbsp;<A href="<%=AppendQuery("../Xbox_Games/Games.asp")%>"><SPAN class=links>Xbox Homepage</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("../DVD/Games.asp")%>"><SPAN class=links>DVD Homepage</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("../Electronics/Games.asp")%>"><SPAN class=links>Electronics Homepage</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("../Applications/Games.asp")%>"><SPAN class=links>Applications Homepage</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("../On_Sale/Games.asp")%>"><SPAN class=links>Items On Sale</SPAN></a><BR>
						</td>
						<td valign=top>
							&#8226;&nbsp;<A href="<%=AppendQuery("NewUser.asp")%>"><SPAN class=links>Become a <i><b>New</b></i> Member</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("LoginTrackOrder.asp")%>"><SPAN class=links>Track an Order</SPAN></a><BR>
							&#8226;&nbsp;<A href="<%=AppendQuery("Viewcart.asp")%>"><SPAN class=links>View my Cart</SPAN></a><BR>			
							&#8226;&nbsp;<A href="<%=AppendQuery("LoginEdit.asp")%>"><SPAN class=links>Edit my User Details</SPAN></a><BR>
						</td>
					</tr>
				</table> 
					 
			</td>
		</tr>
	</table>
	  
	<BR>
</div>