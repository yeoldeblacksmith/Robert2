<!--#include file="../../classes/IncludeList.asp" -->
<%
    dim myUser
    set myUser = new User
    
    myUser.UserName = request.Form("UserName")
    
    myUser.Delete
    
    response.Redirect("default.asp")    
%>