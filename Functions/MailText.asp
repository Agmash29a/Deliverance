<%
'*******************************************************************************************

dim SuccessMail
dim SuccessMailEnd
dim FailMail

SuccessMail = "Deliverance Software Pty Ltd" & CHR(13) & CHR(10) _
			& "ABN 57 096 502 617"  & CHR(13) & CHR(10) & CHR(13) & CHR(10) & CHR(13) & CHR(10) _
			& "Thank you for shopping at Deliverance Software!  When the items you have purchased " _
			& "are ready to be shipped, we will send you another email with an estimated time of " _
			& "arrival and a consignment number for tracking the package."   & CHR(13) & CHR(10) & CHR(13) & CHR(10) _
			& "Listed below are the titles you have purchased:"  & CHR(13) & CHR(10) & CHR(13) & CHR(10) & CHR(13) & CHR(10)
			
SuccessMailEnd = "Again, thanks for shopping at www.deliverance.com.au and if you have any further " _
			   & "questions, please feel free to contact orders@deliverance.com.au"  & CHR(13) & CHR(10) & CHR(13) & CHR(10) & CHR(13) & CHR(10) _
			   & "Kind Regards,"  & CHR(13) & CHR(10) _				
			   & "Orders Team"  & CHR(13) & CHR(10) _
			   & "Deliverance Software"

FailMail = "Deliverance Software Pty Ltd" & CHR(13) & CHR(10) _
		 & "ABN 57 096 502 617"  & CHR(13) & CHR(10) & CHR(13) & CHR(10) & CHR(13) & CHR(10) _	
		 & "Thank you for taking the time to shop at Deliverance Software.  Unfortunately your " _
		 & "credit card transaction has been declined by your issuing bank.  The items you wished " _
		 & "to purchase will remain in your shopping cart for the next thirty days and should you " _
		 & "wish to try again to complete this purchase, we would appreciate doing business with " _
		 & "you. "   & CHR(13) & CHR(10) & CHR(13) & CHR(10) & CHR(13) & CHR(10) _	
		 & "Kind Regards,"  & CHR(13) & CHR(10) _				
		 & "Orders Team"  & CHR(13) & CHR(10) _
		 & "Deliverance Software" 		
		  	  
'*******************************************************************************************
%>
