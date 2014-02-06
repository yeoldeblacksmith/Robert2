<!--#include file="../classes/includelist.asp"-->
<%
    dim myEmail
    set myEmail = new WaiverEmail

    myEmail.SendById request.QueryString(QUERYSTRING_VAR_WAIVERID)

    response.Redirect "emailsent.asp?id=" & request.QueryString(QUERYSTRING_VAR_WAIVERID)
%>