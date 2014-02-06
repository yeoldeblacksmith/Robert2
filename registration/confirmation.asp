<!--#include file="../classes/IncludeList.asp" -->
<%
    dim sUrl : sUrl = SiteInfo.VantoraUrl & "/waiver/default.asp?ev=" & EncodeId(request.Form(PAYPAL_VARIABLE_EVENTID))

    ' send the confirmation email and set the confirmation date/time
    if len(request.Form(PAYPAL_VARIABLE_EVENTID)) > 0 then
        dim myEvent : set myEvent = new ScheduledEvent

        myEvent.Load request.Form(PAYPAL_VARIABLE_EVENTID)
    end if
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Event Registration Confirmation</title>
        <link type="text/css" rel="Stylesheet" href="../content/css/eventregistration.css" />
        <link type="text/css" rel="stylesheet" href="../content/css/jquery.ui.all.css" />
        <link type="text/css" rel="stylesheet" href="../content/css/jquery.mobile-1.1.0.css" />
        <link type="text/css" rel="Stylesheet" href="../content/css/jquery.fancybox-1.3.4.css" />
        <link type="text/css" rel="stylesheet" href="../content/css/waiver.css" />
    </head>

    <body>
        <div id="container" style="margin: 0 auto; text-align: center; width: 740px">
            <div class="waiverSection" style="width: 650px; margin: 15px auto; border: 1px solid grey; text-align:left;">
                <center>
                    <a href="<%= SiteInfo.HomeUrl %>">
                        <img src="<%= LOGO_URL %>" border="0" alt="Go to <%= SiteInfo.Name %> Home Page."  title="Go to <%= SiteInfo.Name %> Home Page.">
                    </a>
                </center>
                <blockquote>
                    <h2 class="center">Step 1 Complete!  Registration Was Successful</h2>
                    <p>You should receive a confirmation email from us.  Check any spam folders if you do not receive it.</p>
                    <h1 align="center">Next Step - Waivers!<br>Parents must sign for their own children!</h1>
                    <p>There are multiple ways to get waivers done from your guests:</p>
                    <ul>
                        <li>Simply forward them this link by email: (copy and paste it) 
                            <br/><br/>
                            <a href="<%= sUrl %>"><%= sUrl %></a>
                            <br/><br/>
                        </li>
                        <li>Or you can tell them to go to 
                            <a href="<%= SiteInfo.HomeUrl %>"><%= SiteInfo.HomeUrl %></a> 
                            and follow the instructions for doing an online waiver.  Make sure your guests have the correct date (<strong><%= myEvent.SelectedDate %></strong>) 
                            and group name (<strong><i><%=myEvent.PartyName %></i></strong>).
                            <br/><br/>
                        </li>
                    </ul>

                    <%= settings(SETTING_BLURB_CONFIRMATION) %>
       
                    <ul>
                        <li>
                            <% if Year(myEvent.SelectedDate) = Year(Now()) then %>
                            You may now <a href="<%= sUrl %>">Sign Your Online Waiver</a>
                            <% else %>
                            Waivers are only good for a calendar year, so make sure you get waivers signed by your participants after Jan 1.
                            <% end if %>
                        </li>
                        <li>
                            <a href="<%= SiteInfo.HomeUrl %>">Go to <%= SiteInfo.Name %></a>
                        </li>
                    </ul>
                </blockquote>
            </div>
        </div>
    </body>
</html>
