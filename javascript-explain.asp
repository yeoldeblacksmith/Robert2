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
       <div align="center"><br><br> <a href="<%= SiteInfo.HomeUrl %>"><img src="<%= LOGO_URL %>" border="0" alt="Go to <%= SiteInfo.Name %> Home Page."  title="Go to <%= SiteInfo.Name %> Home Page."></a><br><br>
      <a href="<%= SiteInfo.VantoraUrl %>/registration"><img src="../../images/booknow.png" border="0"></a>  <a href="<%= SiteInfo.VantoraUrl %>/waiver"><img src="../../images/waiveronline.jpg" border="0"></a> 
      
      <table width="650" border="0" cellspacing="4" cellpadding="4">
  <tr>
    <th colspan="2"><h1>How
        to turn on or &quot;enable&quot; Javascript</h1><br>
        
    </th>
  </tr>
  <tr>
    <td valign="top" nowrap><font size="-1" face="Verdana, Arial, Helvetica, sans-serif"><b>Internet
    Explorer for Windows</b></font></td>
    <td><font size="-1" face="Verdana, Arial, Helvetica, sans-serif">From the &quot;Tools&quot; menu of IE, select the "Internet
        Options". Click on the "Security" tab. Make
      sure the "Internet" globe icon is highlighted. Click on the "Custom Level..." button
      to bring up the security options for your browser. Search through the menu
      for the "Active scripting" option. Make sure "Enable" is selected. Click
      the "OK" button. Close this window and  click the &quot;Refresh (F5)&quot; button of the
      page requiring Javascript.</font></td>
  </tr>
  <tr>
    <td valign="top" nowrap><font size="-1" face="Verdana, Arial, Helvetica, sans-serif"><b>Firefox
        for Windows</b></font></td>
    <td><font size="-1" face="Verdana, Arial, Helvetica, sans-serif">From the "Tools" menu of Firefox, select  &quot;Options...&quot;. Select the "Content" tab and check the &quot;Enable JavaScript&quot; option.
    Click &quot;OK&quot; to close this window, then click the &quot;Reload current page&quot; button of the page requiring Javascript.</font></td>
  </tr>
  <tr>
    <td valign="top" nowrap><font size="-1" face="Verdana, Arial, Helvetica, sans-serif"><b>Safari for Macintosh</b></font></td>
    <td><font size="-1" face="Verdana, Arial, Helvetica, sans-serif">Go to and select  the "Preferences..." option in the &quot;Safari&quot; menu. Click on the &quot;Security&quot; icon in the top row of preference options. Under the &quot;Web Content:&quot; category, make sure the &quot;Enable JavaScript&quot; box is checked. Close this window and  click the &quot;Reload&quot; button of the page requiring Javascript.</font></td>
  </tr>
  <tr>
    <td valign="top" nowrap><font size="-1" face="Verdana, Arial, Helvetica, sans-serif"><b>Safari for Windows</b></font></td>
    <td><font size="-1" face="Verdana, Arial, Helvetica, sans-serif">From the pull down menu that looks like a gear, select  the "Preferences..." option. Click on the &quot;Security&quot; icon in the top row of preference options. Under the &quot;Web Content:&quot; category, make sure the &quot;Enable JavaScript&quot; box is checked. Close this window and  click the &quot;Reload&quot; button of the page requiring Javascript.</font></td>
  </tr>
  <tr>
    <td valign="top" nowrap><font size="-1" face="Verdana, Arial, Helvetica, sans-serif"><b>Firefox for Macintosh</b></font></td>
    <td><font size="-1" face="Verdana, Arial, Helvetica, sans-serif">Choose  "Preferences..." under Firefox's menu bar. Select the "Content" tab and check the &quot;Enable JavaScript&quot; option.
    Close this window and  click the &quot;Reload current page&quot; button of the page requiring Javascript.</font></td>
    </tr>
  
   <tr>
    <td valign="top" nowrap><font size="-1" face="Verdana, Arial, Helvetica, sans-serif"><b>Chrome for PC</b></font></td>
    <td><font size="-1" face="Verdana, Arial, Helvetica, sans-serif">Click on "Options".  
Go to "Under The Hood".  Click on "Content Settings". Click on the "JavaScript" tab. </font></td>
    </tr>  
    
    
</table>

      
      </div>
        
        
        
                </body>
        </html>
        
        