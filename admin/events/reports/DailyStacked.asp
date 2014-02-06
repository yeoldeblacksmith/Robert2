<!DOCTYPE html>
<!--#include file="../../../classes/includelist.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForSiteUsers
%>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title>Daily Report</title>

    <link type="text/css" rel="Stylesheet" href="../../../content/css/eventregistration3.css" />
    <link type="text/css" rel="stylesheet" href="../../../content/css/jquery.ui.all.css" />
    <link type="text/css" rel="Stylesheet" href="../../../content/css/admin.css" />
    <link type="text/css" rel="Stylesheet" href="../../../content/css/navmenu.css" />

        
    <script type="text/javascript" src="https://www.google.com/jsapi"></script>
    <script type="text/javascript" src="../../../content/js/jquery-1.7.2.min.js"></script>
    <script type="text/javascript" src="../../../content/js/jquery.ui.core.min.js"></script>
    <script type="text/javascript" src="../../../content/js/jquery.ui.widget.min.js"></script>
    <script type="text/javascript" src="../../../content/js/jquery.ui.datepicker.min.js"></script>

    <script type="text/javascript">
        $(document).ready(function () {
            var today = new Date;
            var formattedDate = (today.getMonth() + 1) + "/" + today.getDate() + "/" + today.getFullYear();

            var dateField = $("#selectedDate");
            dateField.val(formattedDate);
            dateField.datepicker({
                onSelect: function (dateText, inst) {
                    drawChart(dateText);
                    loadPartyList(dateText);
                }
            });


            dateField.keyup(function (e) {
                if (e.keyCode == 13) {
                    drawChart(dateField.val());
                    loadPartyList(dateField.val());
                    dateField.datepicker('hide');
                }
            });
        });

        google.load('visualization', '1.0', { 'packages': ['corechart'] });
        google.setOnLoadCallback(initialDraw);

        function checkOut(eventId, partyName) {
            var now = new Date();
            var hour = now.getHours();
            var mins = now.getMinutes();

            if (mins < 15)
                mins = 0;
            else if (mins < 45)
                mins = 30;
            else {
                hour += 1;
                mins = 0;
            }

            var hourString = '';
            var minString = '';

            if (hour < 10)
                hourString = '0' + hour;
            else
                hourString = hour;

            if (mins < 10)
                minString = '0' + mins;
            else
                minString = mins;
            
            if(confirm("Do you want to mark the " + partyName + " party as gone?"))
                $.ajax({
                    url: '../../../ajax/event.asp?act=9&id=' + eventId + '&st=' + hourString + ':' + minString,
                    error: function () { alert('Problems encountered checking out'); },
                    success: function () { initialDraw(); }
                });
        }

        function initialDraw() {
            var selectedDate = $("#selectedDate").val();
            drawChart(selectedDate);
            loadPartyList(selectedDate);
        }

        function drawChart(selectedDate) {
            $.ajax({
                url: '../../../ajax/event.asp?act=7&dt=' + selectedDate,
                error: function () { alert("Problems encountered"); },
                success: function (dataArray) {

                    var data = new google.visualization.DataTable(dataArray);
                    var options =
                    {
                        title: "Patrons Scheduled for the Day",
                        width: 874,
                        height: 700,
                        chartArea: { top: 0 },
                        vAxis: { title: "Time" },
                        hAxis: { title: "Number Of Patrons" },
                        isStacked: true
                    };

                    new google.visualization.BarChart(document.getElementById('visualization')).draw(data, options);
                }
            });
        }

        function loadPartyList(selectedDate) {
            $.ajax({
                url: '../../../ajax/event.asp?act=8&dt=' + selectedDate,
                error: function () { alert("Problems encountered"); },
                success: function (htmlData) {
                    var listContainer = $("#list");
                    listContainer.empty();
                    listContainer.html(htmlData);
                }
            });
        }
        
    </script>
</head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Daily Stacked Report"

    navmenu.FormControlStrings.Add "<input type=""text"" id=""selectedDate"" style=""width: 75px"" />"
    navmenu.FormControlStrings.Add "<img src=""../../../content/images/calendar_view_month.png"" alt="""" title=""date"" onClick=""$('#selectedDate').datepicker('show');"" />"

    navmenu.WriteNavigationSection NAVIGATION_NAME_EVENT_SCHEDULE 
%>
        <table class="adminTable">
            <tr>
                <td id="list" />
                <td id="visualization" />
            </tr>
        </table>
    </body>
</html>
