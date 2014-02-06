<%
'Dim mail, body
dim body
body = "Name: " & Request.Form("name") &  "<br>email: " & Request.Form("email") & "<br>Phone: " & Request.Form("phone") & "<br>Field Name: " & Request.Form("fieldname") & "<br>WebSite: " & Request.Form("website") & "<br><br>Notes: <br>" & Request.Form("notes")

'Set mail = Server.CreateObject("CDO.Message")
'mail.To = Request.Form("To")
'mail.From = Request.Form("From")
'mail.Subject = Request.Form("Subject")
'mail.TextBody = body
'mail.Send()

'Set mail = nothing
'Set body = nothing

private mo_emailClient
set mo_emailClient = Server.CreateObject("SMTPsvg.Mailer")

with mo_emailClient
	.FromName = request.QueryString("name")
	.FromAddress = "info@gatsplat.com" 
	.RemoteHost = "64.251.193.6"
	.AddRecipient "Vantora Interest Request", "info@gatsplat.com"
	.Subject = "Admin Password"
	.BodyText = body
	.ContentType = "text/html"

	Send = .SendMail()
end with

response.redirect("thanks.asp")

%>