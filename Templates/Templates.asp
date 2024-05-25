<%









'*------------------------------------------------------------------------------------------*
'*									Feature Tile template                                   *
'*									~~~~~~~~~~~~~~~~~~~~~                                   *
'*																							*
'*																							*
'*------------------------------------------------------------------------------------------*

'**** Follow these steps (5 Steps) ****

'1. Enter the id of the game:
'	eg. FeatureTitleID = 264
	FeatureTitleID = 
	
'2. Enter the filename of the image in the pictures directory
'	eg. FeatureTitleImage = "OpFlashscreen2.jpg"
	FeatureTitleImage = ""

'3. Enter the text that will appear as a heading
'	eg. FeatureTitleHeading = "Feature Title - Operation Flashpoint (PC)"
	FeatureTitleHeading = ""
	
'4. Enter the caption text for the image
'	eg. FeatureTitleCaption = "Where are they hiding?"
	FeatureTitleCaption = ""	
	
%>
<center><a href='<%=AppendQuery("AddToCart.asp?id=" & cstr(FeatureTitleID))%>'><IMG src="../pictures/GreenCheckSmall3.gif" width=10 border=0 ></a>&nbsp;<font class="headingfont"><%=FeatureTitleHeading%></font>
<table><tr height=2><td></td></tr></table><IMG src="../pictures/<%=FeatureTitleImage%>"><br><font class="textfont" size="1"><b><%=FeatureTitleCaption%></b></center><br>
<%

'5. Type the text for the article here (between the two lines)
'	
'<------------- Start text -------------->'
%>

<%
'<------------- End Text -------------->'
%>
</font><br><br>
<%
'*------------------------------------------------------------------------------------------*
'* ---> End of feature title template														*
'*------------------------------------------------------------------------------------------*






















'*------------------------------------------------------------------------------------------*
'*									  Our Picks template                                    *
'*									  ~~~~~~~~~~~~~~~~~~                                    *
'*																							*
'*																							*
'*------------------------------------------------------------------------------------------*

'**** Follow these steps (1 step) ****

'1. Enter the ids of the games:
'	eg. OurPicksIDs = "(23,240,123)"
	OurPicksIDs = "()"
	
%>
<font class="headingfont"><center>Our Picks</center></font><br>
<%
'**** grab all matches

cm.ActiveConnection=Cstr(Application("connstring"))

cm.CommandText = "Select products.boxshoturl, products.id, platform.name platform,products.name,recprice,price,salepercentage,products.ProductDesc from products,platform where products.platformid=platform.id " _
			&    "and products.id in " & OurPicksIDs
								         
set objrsOurPicks = cm.Execute
cm.ActiveConnection=nothing
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
<%	
while not objrsOurPicks.eof
	%>
	<tr>
		<td colspan="2"><center><a style="text-decoration: none" href='<%=AppendQuery("AddToCart.asp?id=" & objrsOurPicks("id"))%>'><IMG src="../pictures/GreenCheckSmall3.gif" width=10 border=0></a>&nbsp;<%=objrsOurPicks("name")%> (<%=objrsOurPicks("platform")%>) - <font  size ="-2">$<%=objrsOurPicks("price")%></font>
		<%
		if  cint(objrsOurPicks("SalePercentage"))>0 then
			%>
			<font color="#BD0000" size ="-2"> On Sale! was $<%=objrsOurPicks("Recprice")%>, now reduced by <%=objrsOurPicks("SalePercentage")%>%</font> 
			<%
		end if
		%>
		</center>
		</td>
	</tr>
	<tr height="1">
		<td colspan="2">&nbsp
		</td>
	</tr>
	<tr>
		<%
		'**** Check for box shot ****
		
		if trim(objrsOurPicks("boxshoturl")) = "" then
			%>
			<td colspan="2"><p align="justify"><font class="textfont"><%=objrsOurPicks("productdesc")%></font></p></td>
			<%
		else
			%>
			<td width="200" valign="top"><a href="individualgame.asp?id=<%=objrsOurPicks("id")%>"><IMG src="<%=objrsOurPicks("boxshoturl")%>" align="top" border=0></a>
			</td>
			<td><p align="justify"><font class="textfont"><%=objrsOurPicks("productdesc")%></font></p></td>
			<%		
		end if
	%>
	</tr>
	<%
	objrsOurPicks.movenext
	
	'**** check for last line ****
	
	if not objrsOurPicks.eof then
		%>
		<tr height="1"><td colspan="2">&nbsp;</td></tr>
		<tr>
			<td  colspan="2"><hr width="90%" size="1" color="#bd0000">
			</td>
		</tr>
		<tr height="1">
			<td colspan="2">&nbsp;
			</td>
		</tr>
		<%
	end if
	
wend	
set objrsOurPicks = nothing
%>
</table><br><br>
<%
'*------------------------------------------------------------------------------------------*
'* ---> End of Our picks template														    *
'*------------------------------------------------------------------------------------------*





























'*------------------------------------------------------------------------------------------*
'*									  Article template                                      *
'*									  ~~~~~~~~~~~~~~~~~                                     *
'*																							*
'*																							*
'*------------------------------------------------------------------------------------------*

'**** Follow these steps (3 Steps) ****

'1. Enter the text for the article heading:
'	eg. ArticleHeading = "The Loud Speaker"
	ArticleHeading = ""
	
'2. Type the text for the article here (between the two lines)
'	if you want to include a gamelink in here use the following	code in step 3
'	remember id goes after --> AppendQuery("AddToCart.asp?id= -->here<-- and the name goes in between the <b> and </b>

'3. delete this entire line once you have finished -->%> <a href="<%=AppendQuery("AddToCart.asp?id=")%>"><IMG border="0" src="../pictures/GreenCheckSmall3.gif" width="10"></a>&nbsp;<b></b> <%

%>
<font class="headingfont"><center><%=ArticleHeading%></center></font><br><font class="textfont" size="1">
<%	
'<------------- Start text -------------->'
%>

<%
'<------------- End Text -------------->'
%>
</font><br><br>
<%
'*------------------------------------------------------------------------------------------*
'* ---> End of Article														                *
'*------------------------------------------------------------------------------------------*























%>