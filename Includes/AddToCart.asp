<!-- #include file="../functions/AppendQuery.asp"-->
<%
Response.Expires=-10000

dim cm				'Command Object
dim objrs			'Recordset

SET cm = Server.CreateObject("ADODB.Command")

IF Request.Cookies("LoggedIn")="TRUE" THEN 

	cm.ActiveConnection=Cstr(Application("connstring"))

	cm.CommandText = "insert into shoppingcart (productid,userid) values (" & request("id") _
	  & "," & Request.Cookies("userid") & ")"
								         
	set objrs = cm.Execute
	cm.ActiveConnection=nothing

	SET cm = NOTHING
	SET objrs = NOTHING

	Response.Redirect("AddedToCart.asp?id=" & Request("id"))

ELSE

	SET cm = NOTHING
	SET objrs = NOTHING
	
	dim rgxp
	dim URL

	'**** first part of URL

	URL = "AddedToCart.asp?id=" & Request("id") 
	
	'**** test for cart string
		
	set rgxp = new regexp

	rgxp.Pattern="cart"

	if not rgxp.test(Request.ServerVariables("QUERY_STRING")) then
	
		URL = URL & "&cart=" & Request("id") & "x"
		
			
	else
	
	URL = replace(URL,"&cart=","")
	
	URL = URL & "&cart=" & Request.QueryString("cart") & Request("id") & "x"

	end if

	Response.Redirect(URL)

END IF
%>
