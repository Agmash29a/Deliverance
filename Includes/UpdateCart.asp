<%
Response.Expires=-10000

dim objrs							   
dim ojbrs2	
dim idtext
dim returntext

SET cm = Server.CreateObject("ADODB.Command")

IF Request.Cookies("LoggedIn")="TRUE" THEN 

	'**** grab all id's in cart of user

	cm.ActiveConnection=Cstr(Application("connstring"))

	cm.CommandText = "select id from shoppingcart where userid=" & Request.Cookies("userid")
								         
	set objrs = cm.Execute
	cm.ActiveConnection=nothing

	'**** if cart item then iterate through them

	if not objrs.EOF then
	
		WHILE not objrs.EOF 

			'**** if selected then delete that item

			idtext = "id" & objrs("id")
			
			if request(idtext)="on" then
			
				cm.ActiveConnection=Cstr(Application("connstring"))
				
				cm.CommandText = "delete from shoppingcart where " _
							   & "id = " & objrs("id")
										         
				set objrs2 = cm.Execute
				cm.ActiveConnection=nothing
				
			end if
		
			objrs.movenext
		
		WEND 
		
		set objrs2 = nothing

	end if
	
else
	
	'**** User not logged in - delete appropriate
	
	dim amatch
	dim matches
	dim rgxp
	dim lastmatch		'**** result of the last match
	dim NewCartText
	dim index			'**** counter for checkboxes
				
	'**** pattern to look for
				
	set rgxp = new regexp
	rgxp.Pattern = "x"
	rgxp.Global = true
				
	lastmatch = 1		'**** virtual x at start of string
	index = 1			'**** first check box
				
	set matches = rgxp.Execute(Request.QueryString("cart"))
				
	for each amatch in matches
	
		idtext = "id" & index
		
		if not Request.form(idtext)="on" then
	
		NewCartText = NewCartText & mid(Request.QueryString("cart"),lastmatch,(amatch.firstindex+1)-(lastmatch))	& "x"	
		
			
		end if	

		index = index + 1
		
		lastmatch = amatch.firstindex+2
	
	next
	
	'**** Assign new string 

END IF

set objrs = nothing
set cm = nothing

if newcarttext="" then

	returntext = "ViewCart.asp?shippingtype=" & request("shippingtype") & "&shipall=" &  request("shipall") 

else

	returntext = "ViewCart.asp?shippingtype=" & request("shippingtype") & "&shipall=" &  request("shipall") & "&cart=" & NewCartText

end if

Response.Redirect(returntext)
%>
