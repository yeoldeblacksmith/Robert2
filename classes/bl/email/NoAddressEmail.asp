<%
class NoAddressEmail
    public sub Send(EmailType, ObjectId)
        dim body

        body = "<p>No Email Address Encountered</p>" & vbCrLf
        body = body & "<p>While sending emails, an object was found that did not have an email address. Details follow:</p>" & vbCrLf
        body = body & "<ul>" & vbCrLf
        body = body & "<li>Site Name: " & SiteInfo.Name & "</li>" & vbCrLf
        body = body & "<li>Site GUID: " & MY_SITE_GUID & "</li>" & vbCrLf
        body = body & "<li>Email Type: " & EmailType & "</li>" & vbCrLf
        body = body & "<li>Object Id: " & ObjectId & "</li>" & vbCrLf
        body = body & "</ul>" & vbCrLf
        
        dim myEmailCon : set myEmailCon = new EmailConnection

        myEmailCon.ServerAddress = settings(SETTING_EMAILSERVER)
        myEmailCon.SenderAddress = settings(SETTING_FROMADDRESS)
        myEmailCon.SenderName = settings(SETTING_FROMNAME)

        myEmailCon.Send SUPPORT_EMAIL_ADDRESS, SUPPORT_EMAIL_ADDRESS, "No email address in Vantora", body
    end sub
end class
%>