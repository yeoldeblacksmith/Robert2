<!--#include file="../classes/includelist.asp"-->
<%
    dim myWaiver
    set myWaiver = new Waiver

    myWaiver.Load request.QueryString(QUERYSTRING_VAR_WAIVERID)

    Response.Expires = 0 
    Response.Buffer = TRUE 
    Response.Clear 
    Response.ContentType = "image/svg+xml"
    Response.Write myWaiver.Signature
%>