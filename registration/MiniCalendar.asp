<!DOCTYPE html>

<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title></title>
        <link type="text/css" rel="Stylesheet" href="../content/css/eventregistration.css" />

        <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                var now = new Date();
                loadCalendar((now.getMonth() + 1) + '/' + now.getDate() + '/' + now.getFullYear());

                var monthControl = $("select[name=month]");
                var yearControl = $("select[name=year]");

                monthControl.val(now.getMonth() + 1);
                yearControl.val(now.getFullYear());
            });

            function loadCalendar(day) {
                $.ajax({
                    url: '../ajax/Calendar.asp?ct=ci&dt=' + day,
                    success: function (data) { $('span[name=calendar]').html(data); },
                    error: function () { alert('problems encountered'); }
                });
            }

            function setDate(day) {
                parent.setSelectedDay(day);
                parent.reloadTime(day);
                parent.$.fancybox.close();
            }
        </script>
    </head>
    <body>

        <select name="month" onchange="loadCalendar(this.value)">
<%
BuildMonthList()
 %>
        </select>

        <span name="calendar"></span>

    </body>
</html>
<%
    sub BuildMonthList()
        dim dt
        dt = cdate(Month(Now) & "/1/" & Year(Now))

        for i = 0 to 3            
		dim currDate
            currDate = DateAdd("m", i, dt)
           response.Write "<option value='" & FormatDateTime(currDate, vbShortDate) & "'>" & MonthName(Month(currDate)) & "</option>"
        next
    end sub
%>