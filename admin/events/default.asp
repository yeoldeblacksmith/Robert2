<!DOCTYPE html>
<!--#include file="../../classes/includelist.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForSiteUsers
%>

<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Vantora - Scheduled Events</title>
        <meta name="viewport" content="width=device-width, initial-scale=1"/>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>

        <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css?<%= ANTI_CACHE_STRING %>" />
        <link type="text/css" rel="stylesheet" href="../../content/css/jquery.ui.all.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/jquery.fancybox-1.3.4.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/jquery.tablesorter.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/admin.css?<%= ANTI_CACHE_STRING %>" />
        <link type="text/css" rel="stylesheet" href="../../content/css/navmenu.css?<%= ANTI_CACHE_STRING %>" />
        <link type="text/css" rel="stylesheet" href="../../content/css/jquery.powertip-light.css" />
        
        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.ui.core.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.ui.widget.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.ui.datepicker.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.tablesorter.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.fancybox-1.3.4.pack.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.powertip.min.js"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                var today = new Date;
                var formattedDate = (today.getMonth() + 1) + "/" + today.getDate() + "/" + today.getFullYear();

                var dateField = $("#selectedDate");
                dateField.val(formattedDate);
                dateField.datepicker({
                    onSelect: function (dateText, inst) { selectDay(dateText); }
                });

                selectDay(formattedDate);

                dateField.keyup(function (e) {
                    if (e.keyCode == 13) {
                        selectDay(dateField.val());
                        dateField.datepicker('hide');
                    }
                });
            });

            function editRegistration(eventId) {
                $.fancybox({
                    'type': 'iframe',
                    'href': 'registration.asp?id=' + eventId + '&vt=' + $("#viewType").val(),
                    'titlePosition': 'outside',
                    'title': '&copy;Vantora',
                    'width': 600,
                    'height': 700,
                    'transitionIn': 'elastic',
                    'transitionOut': 'elastic',
                    'speedIn': 200,
                    'speedOut': 200
                });
            }

            function markAsPaid(Id) {
                $.powerTip.hide();

                if (confirm("Are you sure you want to mark this as paid?")) {
                    $.ajax({
                        url: '../../ajax/Event.asp?act=3&id=' + Id,
                        error: function () { alert("problems encountered while changing status"); },
                        sync: true
                    });

                    reloadData();
                }
            }

            function reloadData() {
                if ($("#viewType").val() == 'day') {
                    selectDay($("#selectedDate").val());
                } else {
                    showAll();
                }
            }

            function resizeFancyBox() {
                $('#fancybox-content').height($('#fancybox-frame').contents().height() + 5);
                $('#fancybox-content').width($('#fancybox-frame').contents().wiidth() + 5);
                $.fancybox.center();
            }

            function selectDay(date) {
                $.ajax({
                    url: '../../ajax/Event.asp?act=1&dt=' + date,
                    success: function (data) {
                        $("#schedule").html(data);
                        $("#scheduleTable").tablesorter();
                        $("#viewType").val("day");
                        $(".tooltip").powerTip({ mouseOnToPopup: true, placement: 'se', smartPlacement: true });
                    },
                    error: function () { alert("problems encountered"); }
                });
            }

            function showAll() {
                $.ajax({
                    url: '../../ajax/Event.asp?act=2',
                    success: function (data) {
                        $("#schedule").html(data);
                        $("#scheduleTable").tablesorter();
                        $("#viewType").val("all");
                        $(".tooltip").powerTip({ mouseOnToPopup: true, placement: 'se', smartPlacement: true });
                    },
                    error: function () { alert("problems encountered"); }
                });
            }
        </script>
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Registration Calendar"

    navmenu.FormControlStrings.Add "<input type=""text"" id=""selectedDate"" style=""width: 75px"" tabindex=""1"" />"
    navmenu.FormControlStrings.Add "<img src=""../../content/images/calendar_view_month.png"" alt="""" title=""Selected Date"" onclick=""$('#selectedDate').datepicker('show');"" tabindex=""2"" />"
    navmenu.FormControlStrings.Add "<img src=""../../content/images/calendar.png"" alt="""" title=""Show All"" onclick=""showAll();"" tabindex=""3"" />"

    navmenu.WriteNavigationSection NAVIGATION_NAME_EVENT_SCHEDULE 
%>
        <input type="hidden" id="viewType" value="day" />

        <p></p>

        <span id="schedule"></span>

        <br /><br /><br />
    </body>
</html>