<%
function AppendQuery(OriginalURL)

'**** Append querystring with quids or crtid's
	
	'**** test for a ?
	
	dim rgxp		'regexp objext
	dim v			'query strings
	dim changes		'was the url changed/
	dim URL			'URL to jump to
	
	URL = OriginalURL
	
	set rgxp = new regexp

	rgxp.Pattern="\?"

	if not rgxp.test(URL) then
	
		URL = URL & "?"
		
	else
	
		'**** append an &
	
		URL = URL & "&"
	
	end if

	'**** append guid's and cart id's to querystring	

	for each v in Request.QueryString
	
		if v = "guid" then
		
			URL = URL & "guid=" & Request.QueryString(v) & "&"
		
			changes = true
			
		end if
	
		if v = "cart" then
		
			URL = URL & "cart=" & Request.QueryString(v) & "&" 
			
			changes = true
		
		end if
	
	next
	
	'**** Chop off last &
	
	if right(URL,1) = "&" or right(URL,1) = "?" then
	
		URL = left(URL,(len(URL)-1))

	end if

	if changes then
	
		AppendQuery = URL 

	else
	
		AppendQuery = OriginalURL	
	
	end if

end function
%>