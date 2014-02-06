<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Event Registration</title>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        
        <link type="text/css" rel="Stylesheet" href="content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="content/css/jquery.fancybox-1.3.4.css" />

        <script type="text/javascript" src="content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="content/js/jquery.fancybox-1.3.4.pack.js"></script>
        <script type="text/javascript">
            function showRegistration(selectedDate) {
                $.fancybox({
                    'type': 'iframe',
                    'href': 'EventRegistration.asp?dt=' + selectedDate,
                    'titlePosition': 'outside',
                    'title': '&copy;GatSplat',
                    'width': 600,
                    'height': 550,
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
                    'height': 350,
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
        
        <h1 style="text-align: center">Event Registration</h1>
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
        <br /><br />
        <h2 align="center">Instructions</h2> 
        <ul>
            <li>Choose the Month from the Drop Down selector.</li>
            <li>Find the date you want to book for.</li>
            <li>If the date is white, it means that it is not one of the days we are open.  You can still book private events for those days, but your would need to <a href="http://www.gatsplat.com/contact.asp">call us</a> to set up your private event.</li>
            <li>You will notice the times we are open are listed on each day.  You will be able to set reservations from opening, up until 2 hours before we close.</li>
            <li>Click on the Date you want to book your event for and fill out the booking form.</li>
        </ul>        
        	
        
        
        
        </td></tr></table>
        <br /><br />
       
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