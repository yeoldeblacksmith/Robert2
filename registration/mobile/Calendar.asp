<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Event Registration</title>
        <meta name="viewport" content="width=device-width, initial-scale=1">
		<meta name="apple-mobile-web-app-capable" content="yes" />
	    
        <link rel="stylesheet" href="../content/css/jquery.mobile-1.1.0.css" />
        <link rel="Stylesheet" href="../content/css/registrationmobile.css" />
		
        <style type="text/css">
			.center {
				text-align: center;
			}
			.center * {
				margin: 0 auto;
			}
		</style>

        <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
		<script type="text/javascript" src="../content/js/jquery.mobile-1.1.0.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.cookie.js"></script>
		<script type="text/javascript" src="../content/js/common.js"></script>
        <script type="text/javascript">
            function loadCalendar(day) {
                $.ajax({
                    url: '../common/CalendarAjax.asp?ct=cm&dt=' + day,
                    success: function (data) {
                        $('#calendar').html(data);
                        setCellHeight(); 
                    },
                    error: function () { alert('problems encountered'); }
                });
            }

            $(document).live('pageinit', function () {
                var now = new Date();
                loadCalendar((now.getMonth() + 1) + '/' + now.getDate() + '/' + now.getFullYear());

                var monthControl = $("select[name=month]");
                monthControl.val(now.getMonth() + 1);
            });

            function setDate(day) {
                $("#selectedDate").val(day);
                document.forms[0].submit();
            }    
        </script>
    </head>
    
    <body>
        <form action="Registration.asp" method="post">
            <input type="hidden" id="selectedDate" name="selectedDate" value="" />
        </form>

        <!-- Start of first page: #home -->
		<div data-role="page" id="home" data-theme="b">
            <!-- header -->
            <div data-role="header">
				<a href ="http://www.gatsplat.com/mobile/index.asp" rel="external" data-role="button" data-icon="home">Home</a><h1>Gatsplat</h1>
			</div>
			
            <!-- logo -->
            <div class="center" style="background: #ffffff; box-shadow: 0px 7px 5px #2d2d2d;">
                <img style="max-height: 200px; " src="../content/images/mobile/logo.png" alt="Gatsplat logo"/>
			</div>
			
            <!-- content -->
            <div data-role="content">
                <select name="month" onchange="loadCalendar(this.value)">
<%
BuildMonthList()
 %>
                </select>

                <p />

                <span id="calendar"></span>
            </div>

            <!-- footer -->
            <div data-role="footer" class="center" style="height: 55px;">
				<h4 style="position: relative; bottom: 8px;">
                    <span class="social_up">Like us </span>
                    <a href="http://m.facebook.com/gatsplat"><img src="../content/images/mobile/facebook.png" alt="facebook icon" /></a>
                    <a href="http://mobile.twitter.com/gatsplat"><img src="../content/images/mobile/twitter.png" alt="twitter icon"/></a>
                    <span class="social_up"> Follow us </span>
                </h4>
			</div>
        </div>
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