<%
'*******************************************************************************************

SUB SendEmail(MailFrom, MailTo, MailSubject, MailBody)

'*-----------------------*
'* Send an email via CDO *
'*-----------------------*

dim MyMail

Set MyMail = Server.CreateObject("CDONTS.NewMail")
MyMail.From = MailFrom
MyMail.To = MailTo
MyMail.Bcc = "orders@deliverance.com.au"
MyMail.Subject = MailSubject
MyMail.Body = MailBody
MyMail.Importance = 1
MyMail.Send
set MyMail = nothing
	
END SUB	

'*******************************************************************************************
%>
