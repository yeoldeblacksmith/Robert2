<!--#include file="../classes/IncludeList.asp"-->
<%
    'dim sUrl : sUrl = SiteInfo.VantoraUrl & "/waiver/default.asp?ev=" & request.Form(QUERYSTRING_VAR_ID)
    dim sUrl : sUrl = SiteInfo.VantoraUrl & "/waiver/default.asp?ev=" & request.Form("EventId")

    ' send the confirmation email and set the confirmation date/time
    'if len(request.Form(QUERYSTRING_VAR_ID)) > 0 then
    if len(request.Form("EventId")) > 0 then
        dim myEvent : set myEvent = new ScheduledEvent

        'myEvent.Load DecodeId(request.Form(QUERYSTRING_VAR_ID))
        myEvent.Load DecodeId(request.Form("EventId"))
    end if
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Payment Received</title>
        <link type="text/css" rel="stylesheet" href="../content/css/jquery.mobile-1.1.0.css" />
        <link type="text/css" rel="Stylesheet" href="../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../content/css/waiver.css" />

        <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.mobile-1.1.0.min.js"></script>
    </head>
    <body>
        <div id="container" style="margin: 0 auto; text-align: center; width: 740px">
            <div class="waiverSection" style="width: 650px; margin: 15px auto; border: 1px solid grey; text-align:left;">
                <center>
                    <a href="<%= SiteInfo.HomeUrl %>">
                        <img src="<%= LOGO_URL %>" border="0" alt="Go to <%= SiteInfo.Name %> Home Page."  title="Go to <%= SiteInfo.Name %> Home Page.">
                    </a>
                </center>
            
                <h3 class="center">Payment Received</h3>
                <p class="center">Thank you! Your payment has been received.</p>

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
