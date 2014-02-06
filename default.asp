<!DOCTYPE html>
<!--#include file="classes/includelist.asp"-->
<%
    dim eventId, mySite
    set mySite = new Site
    mySite.Load MY_SITE_GUID

    if len(Request.QueryString(QUERYSTRING_VAR_EVENTID)) = 0 then
        eventId = ""
    else
        eventId = DecodeId(Request.QueryString(QUERYSTRING_VAR_EVENTID))
    end if
%>

<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title><%= SiteInfo.Name %> </title>
        
    </head>
    <body>
        <div align="center">
            <br/><br/>
            <a href="<%= SiteInfo.HomeUrl %>">
                <img src="<%= LOGO_URL %>" border="0" alt="Go to <%= SiteInfo.Name %> Home Page."  title="Go to <%= SiteInfo.Name %> Home Page.">
            </a>
            <br/><br/>
            <a href="<%= SiteInfo.VantoraUrl %>/registration">
                <img src="../../images/booknow.png" border="0"/>
            </a> 
             <a href="<%= SiteInfo.VantoraUrl %>/waiver">
                 <img src="../../images/waiveronline.jpg" border="0"/>
             </a>
        </div>
    </body>
</html>
        
        