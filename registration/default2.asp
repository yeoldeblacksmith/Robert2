<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html>
    <head>
        <title>Event Registration</title>
        <meta name="description" content="Indoor Paintball fields perfect for birthday parties, bachelor parties as well as corporate team building events or any event in DFW." />
        <meta Name="keywords" Content="paintball games, dfw, venue, dallas, paintball field, indoor paintball, paintball supplies, gatsplat, birthday parties,  ">
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <!--#include file="../nghead.asp"-->
        <!--#include file="../countusers/UPDATER.ASP" -->

        <meta name="google-site-verification" content="IUE9L0Evw9Ovgc86mXY-AMjCjyFXteKGTp0NiB6eE7U" />

        <link type="text/css" rel="Stylesheet" href="content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="content/css/jquery.fancybox-1.3.4.css" />

        <script type="text/javascript" src="content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="content/js/jquery.fancybox-1.3.4.pack.js"></script>
        <script type="text/javascript">
            function showGuestList(eventId) {
                $.fancybox({
                    'type': 'iframe',
                    'href': 'SimpleInvite.asp?id=' + eventId,
                    'titlePosition': 'outside',
                    'title': '&copy;GatSplat',
                    'width': 600,
                    'height': 650,
                    'transitionIn': 'elastic',
                    'transitionOut': 'elastic',
                    'speedIn': 200,
                    'speedOut': 200
                });
            }

            function showRegistration(selectedDate) {
                $.fancybox({
                    'type': 'iframe',
                    'href': 'EventRegistration2.asp?dt=' + selectedDate,
                    'titlePosition': 'outside',
                    'title': '&copy;GatSplat',
                    'width': 600,
                    'height': 700,
                    'transitionIn': 'elastic',
                    'transitionOut': 'elastic',
                    'speedIn': 200,
                    'speedOut': 200
                });
            }

            function showSent() {
                $.fancybox({
                    'type': 'iframe',
                    'href': 'InvitesSent.htm',
                    'titlePosition': 'outside',
                    'title': '&copy;GatSplat',
                    'width': 500,
                    'height': 450,
                    'transitionIn': 'elastic',
                    'transitionOut': 'elastic',
                    'speedIn': 200,
                    'speedOut': 200
                });
            }

            function showSuccess() {
                $.fancybox({
                    'type': 'iframe',
                    'href': 'EventSubmitted.htm',
                    'titlePosition': 'outside',
                    'title': '&copy;GatSplat',
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
                    url: 'common/CalendarAjax.asp?ct=cu&dt=' + day,
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
        <!--#include file="..\gatmenuhead.asp"--></blockquote>
        <h1 style="text-align: center">Days, Hours & Event Registration</h1>
<table><tr><td width="600">

        <span>


            <span style="padding: 0 15px 0 30px">Month:</span>
            <select name="month" onchange="loadCalendar(this.value)">
<%
BuildMonthList()
 %>
            </select>
        </span>

        <span name="calendar"></span>
        </td><td>&nbsp;&nbsp;</td><td valign="top">
        <h2 align="center">Instructions</h2> 
        <ul>
            <li>Choose the Month from the Drop Down selector.</li>
            <li>The Days we are open are in Green, along with the hours we are open.
        </ul>
        <h2 align="center">To Make a Reservation</h2>
        <ul>
            <li>If you want to reserve a day and time, just click on the Day of your Choice. 
            <li>If the date is white, it means that it is not one of the days we are open.  You can still book private events for those days, but your would need to <a href="http://www.gatsplat.com/contact.asp">call us</a> to set up your private event.</li>
            <li>You will notice the times we are open are listed on each day.  You will be able to set reservations from opening, up until 2 hours before we close.</li>
            <li>Click on the Date you want to book your event for and fill out the booking form.</li>
        </ul>    
        
        
        
        </td></tr></table>
        <br /><br />
        <!--#include file="../gatsplatfooter5.asp"-->
    </body>
</html>

<%
    sub BuildMonthList()
        dim dt
        dt = cdate(Month(Now) & "/1/" & Year(Now))

        for i = 0 to 3            
		dim currDate
            currDate = DateAdd("m", i, dt)
           response.Write "<option value='" & FormatDateTime(currDate, vbShortDate) & "'>" & MonthName(MOnth(currDate)) & "</option>"
        next
    end sub
%>