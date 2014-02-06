<%
class ExceptionEmail
    public sub Send(oRequest, sErrNumber, sErrDescription, sErrSource, oAspError)
        'dim objASPError : Set objASPError = Server.GetLastError
        dim body

        body = "<p>Unhandled exception encountered:</p>"

        body = body & "<h2>Vantora Site Info:</h2>" & vbCrLf
        body = body & "<ul>" & vbCrLf
        body = body & "<li>Site Name: " & SiteInfo.Name & "</li>" & vbCrLf
        body = body & "<li>Site GUID: " & MY_SITE_gUID & "</li>" & vbCrLf
        body = body & "</ul>" & vbCrLf

        body = body & "<h2>ASP Error Info:</h2>" & vbCrLf
        body = body & "<ul>" & vbCrLf
        body = body & "<li>Err.Number: " & sErrNumber & "</li>" & vbCrLf
        body = body & "<li>Err.Description: " & sErrDescription & "</li>" & vbCrLf
        body = body & "<li>Err.Source: " & sErrSource & "</li>" & vbCrLf
        body = body & "<li>ASPCode: " & Server.HTMLEncode(oAspError.ASPCode) & "</li>" & vbCrLf
        body = body & "<li>Number: 0x" & Hex(oAspError.Number) & "</li>" & vbCrLf
        body = body & "<li>Source: [" & Server.HTMLEncode(oAspError.Source) & "]" & "</li>" & vbCrLf
        body = body & "<li>Category: " & Server.HTMLEncode(oAspError.Category) & "</li>" & vbCrLf
        body = body & "<li>File: " & Server.HTMLEncode(oAspError.File) & "</li>" & vbCrLf
        body = body & "<li>Line: " & CStr(oAspError.Line) & "</li>" & vbCrLf
        body = body & "<li>Column: " & CStr(oAspError.Column) & "</li>" & vbCrLf
        body = body & "<li>Description: " & Server.HTMLEncode(oAspError.Description) & "</li>" & vbCrLf
        body = body & "<li>ASP Description: " & Server.HTMLEncode(oAspError.ASPDescription) & "</li>" & vbCrLf
        'body = body & "<li>Server Variables: " & vbCrLf & Server.HTMLEncode(Request.ServerVariables("ALL_HTTP")) & vbCrLf
        'body = body & "<li>QueryString: " & Server.HTMLEncode(Request.QueryString) & vbCrLf
        'body = body & "<li>URL: " & Server.HTMLEncode(Request.ServerVariables("URL")) & vbCrLf
        'body = body & "<li>Content Type: " & Server.HTMLEncode(Request.ServerVariables("CONTENT_TYPE")) & vbCrLf
        'body = body & "<li>Content Length: " & Server.HTMLEncode(Request.ServerVariables("CONTENT_LENGTH")) & vbCrLf
        'body = body & "<li>Local Addr: " & Server.HTMLEncode(Request.ServerVariables("LOCAL_ADDR")) & vbCrLf
        'body = body & "<li>Remote Addr: " & Server.HTMLEncode(Request.ServerVariables("LOCAL_ADDR")) & vbCrLf
        body = body & "<li>Time: " & Now & vbCrLf
        body = body & "</ul>" & vbCrLf
        
        body = body & "<h2>Form Collection:</h2>" & vbCrLf
        body = body & "<ul>" & vbCrLf
        for each formKey in oRequest.Form
            body = body & "<li>" & formKey & ": " & oRequest.Form(formKey) & "</li>" & vbCrLf
        next    
        body = body & "</ul>" & vbCrLf
        
        body = body & "<h2>QueryString Collection:</h2>" & vbCrLf
        body = body & "<ul>" & vbCrLf
        for each qsKey in oRequest.QueryString
            body = body & "<li>" & qsKey & ": " & oRequest.QueryString(qsKey) & "</li>" & vbCrLf
        next    
        body = body & "</ul>" & vbCrLf

        body = body & "<h2>Server Variable Collection:</h2>" & vbCrLf
        body = body & "<ul>" & vbCrLf
        for each svKey in oRequest.ServerVariables
            body = body & "<li>" & svKey & ": " & oRequest.ServerVariables(svKey) & "</li>" & vbCrLf
        next    
        body = body & "</ul>" & vbCrLf

        dim myEmailCon : set myEmailCon = new EmailConnection

        myEmailCon.ServerAddress = settings(SETTING_EMAILSERVER)
        myEmailCon.SenderAddress = settings(SETTING_FROMADDRESS)
        myEmailCon.SenderName = settings(SETTING_FROMNAME)

        myEmailCon.Send SUPPORT_EMAIL_ADDRESS, SUPPORT_EMAIL_ADDRESS, "Unhandled Exception in Vantora", body
    end sub
end class
%>