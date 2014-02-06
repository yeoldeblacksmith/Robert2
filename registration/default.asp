<!DOCTYPE html>
<!--#include file="../classes/IncludeList.asp" -->
<%
    if cbool(settings(SETTING_MODULE_REGISTRATION)) = false then response.Redirect("disabled.asp")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title><%= SiteInfo.Name %> On Line Registration</title>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">

        <link type="text/css" rel="Stylesheet" href="../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../content/css/jquery.fancybox-1.3.4.css" />
        <link type="text/css" rel="Stylesheet" href="../content/css/waiver.css" />

        <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.fancybox-1.3.4.pack.js"></script>
        <script type="text/javascript">
            function showRegistration(selectedDate) {
                window.location.href = 'registration.asp?dt=' + selectedDate;
            }

            function showSuccess() {
                $.fancybox({
                    'type': 'iframe',
                    'href': 'EventSubmitted.htm',
                    'titlePosition': 'outside',
                    'title': '&copy;<%= SiteInfo.Name %>',
                    'width': 500,
                    'height': 450,
                    'transitionIn': 'elastic',
                    'transitionOut': 'elastic',
                    'speedIn': 200,
                    'speedOut': 200
                });
            }

            function loadCalendar(day) {
                $.ajax({
                    url: '../ajax/Calendar.asp?ct=cu&dt=' + day,
                    success: function (data) { $('span[name=calendar]').html(data); },
                    error: function () { alert('problems encountered'); }
                });
            }

            $(document).ready(function () {
                var now = new Date();
                loadCalendar((now.getMonth() + 1) + '/' + now.getDate() + '/' + now.getFullYear());

                var monthControl = $("select[name=month]");
                var yearControl = $("select[name=year]");

                monthControl.val(now.getMonth() + 1);
                yearControl.val(now.getFullYear());
            });
        </script>
    </head>

    <body>
    
        <div class="waiverSection blockCenter" style="width: 1000px">
            <img src="<%= LOGO_URL %>" alt="<%= SiteInfo.Name %>" class="blockCenter" style="display: block;">
            <h1 style="text-align: center">Days, Hours & Event Registration</h1>

            <table>
                <tr>
                    <td style="width: 600px">
                        <span>
                            <span style="padding: 0 15px 0 30px">Month:</span>
                            <select name="month" onChange="loadCalendar(this.value)">
                                <% BuildMonthList() %>
                            </select>
                        </span>
                        <span name="calendar"></span>
                    </td>
                    <td>&nbsp;&nbsp;</td>
                    <td style="vertical-align: top">
                        <h2 style="text-align: center">Instructions</h2> 
                        <ul>
                            <li>Choose the Month from the Drop Down selector.</li>
                            <li>The Days we are open are in Green, along with the hours we are open.
                        </ul>
                        <h2 style="text-align: center">To Make a Reservation</h2>
                        <ul>
                            <li>If you want to reserve a day and time, just click on the Day of your Choice. 
                            <li>If the date is white, it means that it is not one of the days we are open.  You can still book private events for those days, but your would need to call us: <b><%= SiteInfo.PhoneNumber %></b> to set up your private event.</li>
                            <li>You will notice the times we are open are listed on each day.  You will be able to set reservations from opening, up until 2 hours before we close.</li>
                            <li>Click on the Date you want to book your event for and fill out the booking form.</li>
                        </ul>    
                        <%= Settings(SETTING_BLURB_CALENDAR) %>
                    </td>
                </tr>
            </table>
            
            <% if UsesPayments() then %>
            <!-- PayPal Logo -->
            <table border="0" cellpadding="10" cellspacing="0" align="center" style="margin: 0 auto">
                <tr>
                    <td align="center"></td>
                </tr>
                <tr>
                    <td align="center">
                        <a href="https://www.paypal.com/webapps/mpp/paypal-popup" title="How PayPal Works" onclick="javascript:window.open('https://www.paypal.com/webapps/mpp/paypal-popup','WIPaypal','toolbar=no, location=no, directories=no, status=no, menubar=no, scrollbars=yes, resizable=yes, width=1060, height=700'); return false;">
                            <img src="https://www.paypalobjects.com/webstatic/mktg/logo/AM_SbyPP_mc_vs_dc_ae.jpg" border="0" alt="PayPal Acceptance Mark" style="display: none">
                            <img src="https://www.vantora.com/images/securepaypal.gif" border="0" alt="PayPal Acceptance Mark">
                        </a>
                    </td>
                </tr>
            </table>
            <!-- PayPal Logo -->
            <% end if %>
            
            <p style="text-align:center">
                <a href="https://www.vantora.com" target="_blank"> Online Reservations and Waivers </a> provided by <a href="https://www.vantora.com">Vantora.com</a>
            </p>
        </div>
    </body>
</html>

<%
    sub BuildMonthList()
        dim dt
        dt = cdate(Month(Now) & "/1/" & Year(Now))

        for i = 0 to CInt(Settings(SETTING_REGISTRATION_DISPLAYMONTHS)) - 1
		dim currDate
            currDate = DateAdd("m", i, dt)
           response.Write "<option value='" & FormatDateTime(currDate, vbShortDate) & "'>" & MonthName(MOnth(currDate)) & "</option>"
        next
    end sub
%>