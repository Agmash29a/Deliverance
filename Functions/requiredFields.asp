<%
dim ToEmattersString
dim ToEmattersStringUsr

ToEmattersString = "https://merchant.ematters.com.au/cmaonline.nfs/ePayForm?OpenForm&amp;Seq=1?CompanyName=Deliverance&ReturneMail=" & server.urlencode("orders@deliverance.com.au") & "&" _
				 & "ReturnHttp=" & server.urlencode("https://secure.deliverance.com.au/admin/Decrypt.asp")  & "&MerchantID=291&Sendemail=No&ABN=ABN57096502617&" _
				 & "Bank=BankSA&Platform=ASP&Mode=Test&readers=SYD0291&__Click=0"

ToEmattersStringUsr = "https://merchant.ematters.com.au/cmaonline.nfs/ePayForm?OpenForm&amp;Seq=1?CompanyName=Deliverance&ReturneMail=" & server.urlencode("orders@deliverance.com.au") & "&" _
				 & "ReturnHttp=" & server.urlencode("https://secure.deliverance.com.au/admin/Decrypt.asp")  & "&MerchantID=291&Sendemail=No&ABN=ABN57096502617&" _
				 & "Bank=BankSA&Platform=ASP&Mode=Test&readers=SYD0291&__Click=0"
%>
