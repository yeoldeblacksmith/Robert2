<!DOCTYPE html>
<!--#include file="../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForAdminOnly

    dim bSaved
    bSaved = false

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        with SiteInfo
            .MondayOpenTime = Request.Form("txtMondayOpen")
            .MondayCloseTime = Request.Form("txtMondayClose")
            .TuesdayOpenTime = Request.Form("txtTuesdayOpen")
            .TuesdayCloseTime = Request.Form("txtTuesdayClose")
            .WednesdayOpenTime = Request.Form("txtWednesdayOpen")
            .WednesdayCloseTime = Request.Form("txtWednesdayClose")
            .ThursdayOpenTime = Request.Form("txtThursdayOpen")
            .ThursdayCloseTime = Request.Form("txtThursdayClose")
            .FridayOpenTime = Request.Form("txtFridayOpen")
            .FridayCloseTime = Request.Form("txtFridayClose")
            .SaturdayOpenTime = Request.Form("txtSaturdayOpen")
            .SaturdayCloseTime = Request.Form("txtSaturdayClose")
            .SundayOpenTime = Request.Form("txtSundayOpen")
            .SundayCloseTime = Request.Form("txtSundayClose")
            
            .Save

            bSaved = true
        end with
    end if
%>

<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Available Dates</title>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">

        <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/jquery.fancybox-1.3.4.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/admin.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/navmenu.css" />

        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.fancybox-1.3.4.pack.js"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                var now = new Date();
                loadCalendar((now.getMonth() + 1) + '/' + now.getDate() + '/' + now.getFullYear());

                var monthControl = $("#month");
                var yearControl = $("#year");

                monthControl.val(now.getMonth() + 1);
                yearControl.val(now.getFullYear());

                monthControl.change(function () { selectedDateChanged(); });
                yearControl.change(function () { selectedDateChanged(); });
            });

            function changeOpenHours(day) {
                $.fancybox({
                    'type': 'iframe',
                    'href': 'ChangeOpenHours.asp?dt=' + day,
                    'titlePosition': 'outside',
                    'title': '&copy;Vantora',
                    'width': 400,
                    'height': 160,
                    'transitionIn': 'elastic',
                    'transitionOut': 'elastic',
                    'speedIn': 200,
                    'speedOut': 200,
                    'hideOnOverlayClick': false,
                    'showCloseButton': false
                });
            }

            function loadCalendar(day) {
                $.ajax({
                    url:        '../../ajax/Calendar.asp?ct=ca&dt=' + day,
                    success:    function(data) { $('#calendar').html(data); }, 
                    error:      function () { alert('problems encountered'); }
                });
            }

            function selectedDateChanged() {
                var month = $("#month").val();
                var day = 1;
                var year = $("#year").val();

                loadCalendar(month + '/' + day + '/' + year);
            }

            function validateForm() {
                var valid = true;

                var monOpen = $("#txtMondayOpen");
                var monClose = $("#txtMondayClose");
                var tueOpen = $("#txtTuesdayOpen");
                var tueClose = $("#txtTuesdayClose");
                var wedOpen = $("#txtWednesdayOpen");
                var wedClose = $("#txtWednesdayClose");
                var thuOpen = $("#txtThursdayOpen");
                var thuClose = $("#txtThursdayClose");
                var friOpen = $("#txtFridayOpen");
                var friClose = $("#txtFridayClose");
                var satOpen = $("#txtSaturdayOpen");
                var satClose = $("#txtSaturdayClose");
                var sunOpen = $("#txtSundayOpen");
                var sunClose = $("#txtSundayClose");

                if (!validateTime(monOpen.val())) {
                    valid = false;
                    monOpen.css("border", "1px solid red");
                } else {
                    monOpen.css("border", "");
                }
                if (!validateTime(monClose.val())) {
                    valid = false;
                    monClose.css("border", "1px solid red");
                } else {
                    monClose.css("border", "");
                }
                if (!validateTime(tueOpen.val())) {
                    valid = false;
                    tueOpen.css("border", "1px solid red");
                } else {
                    tueOpen.css("border", "");
                }
                if (!validateTime(tueClose.val())) {
                    valid = false;
                    tueClose.css("border", "1px solid red");
                } else {
                    tueClose.css("border", "");
                }
                if (!validateTime(wedOpen.val())) {
                    valid = false;
                    wedOpen.css("border", "1px solid red");
                } else {
                    wedOpen.css("border", "");
                }
                if (!validateTime(wedClose.val())) {
                    valid = false;
                    wedClose.css("border", "1px solid red");
                } else {
                    wedClose.css("border", "");
                }
                if (!validateTime(thuOpen.val())) {
                    valid = false;
                    thuOpen.css("border", "1px solid red");
                } else {
                    thuOpen.css("border", "");
                }
                if (!validateTime(thuClose.val())) {
                    valid = false;
                    thuClose.css("border", "1px solid red");
                } else {
                    thuClose.css("border", "");
                }
                if (!validateTime(friOpen.val())) {
                    valid = false;
                    friOpen.css("border", "1px solid red");
                } else {
                    friOpen.css("border", "");
                }
                if (!validateTime(friClose.val())) {
                    valid = false;
                    friClose.css("border", "1px solid red");
                } else {
                    friClose.css("border", "");
                }
                if (!validateTime(satOpen.val())) {
                    valid = false;
                    satOpen.css("border", "1px solid red");
                } else {
                    satOpen.css("border", "");
                }
                if (!validateTime(satClose.val())) {
                    valid = false;
                    satClose.css("border", "1px solid red");
                } else {
                    satClose.css("border", "");
                }
                if (!validateTime(sunOpen.val())) {
                    valid = false;
                    sunOpen.css("border", "1px solid red");
                } else {
                    sunOpen.css("border", "");
                }
                if (!validateTime(sunClose.val())) {
                    valid = false;
                    sunClose.css("border", "1px solid red");
                } else {
                    sunClose.css("border", "");
                }

                if (!valid) {
                    $("#valMessage").html("please use military time format (00:00).");
                    $("#valMessage").css("display", "block");
                } else {
                    $("#valMessage").css("display", "none");
                }

                return valid;
            }

            function validateTime(fieldValue) {
                var timeRegEx = /^[012]?[0-9]:[0-9][0-9]$/;
                var valid = true;

                if (!timeRegEx.test(fieldValue)) {
                    valid = false;
                }

                return valid;
            }
        </script>
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Available Dates"

    navmenu.WriteNavigationSection NAVIGATION_NAME_SETTINGS 
%>
        <form method="post">
            <table class="adminTable" style="margin: 0 auto; width: 660px; height: auto;">
                <tr>
                    <td class="adminCell" style=" padding: 10px">
                        <div class="center">
                            Month:&nbsp;
                            <select id="month" onchange="">
                                <option value="1">January</option>
                                <option value="2">February</option>
                                <option value="3">March</option>
                                <option value="4">April</option>
                                <option value="5">May</option>
                                <option value="6">June</option>
                                <option value="7">July</option>
                                <option value="8">August</option>
                                <option value="9">September</option>
                                <option value="10">October</option>
                                <option value="11">November</option>
                                <option value="12">December</option>
                            </select>
                            &nbsp;
                            Year:
                            &nbsp;
                            <select id="year">
            <% 
                Response.Write "<option value=""" & Year(Now()) & """>" & Year(Now()) & "</option>"
                Response.Write "<option value=""" & Year(Now()) + 1 & """>" & Year(Now()) + 1 & "</option>"
            %>
                            </select>
                        </div>
        
                        <div id="calendar"></div>
                    </td>
                </tr>

                <tr>
                    <td class="adminCell" style="padding: 10px">
                        <p class="center">This page allows you to set open and close times for your specific installation/location. These are the times that events are allowed to be scheduled by default.</p>
                    
                        <table class="tableDefault marginCenter noBorder">
                            <tr>
                                <th/>
                                <th>Mon</th>
                                <th>Tue</th>
                                <th>Wed</th>
                                <th>Thu</th>
                                <th>Fri</th>
                                <th>Sat</th>
                                <th>Sun</th>
                            </tr>
                            <tr>
                                <th>Open</th>
                                <td>
                                    <input type="text" name="txtMondayOpen" id="txtMondayOpen" style="width: 50px" maxlength="5" value="<%= GetMilitaryTime(SiteInfo.MondayOpenTime) %>" />
                                </td>
                                <td>
                                    <input type="text" name="txtTuesdayOpen" id="txtTuesdayOpen" style="width: 50px" maxlength="5" value="<%= GetMilitaryTime(SiteInfo.TuesdayOpenTime) %>" />
                                </td>
                                <td>
                                    <input type="text" name="txtWednesdayOpen" id="txtWednesdayOpen" style="width: 50px" maxlength="5" value="<%= GetMilitaryTime(SiteInfo.WednesdayOpenTime) %>" />
                                </td>
                                <td>
                                    <input type="text" name="txtThursdayOpen" id="txtThursdayOpen" style="width: 50px" maxlength="5" value="<%= GetMilitaryTime(SiteInfo.ThursdayOpenTime) %>" />
                                </td>
                                <td>
                                    <input type="text" name="txtFridayOpen" id="txtFridayOpen" style="width: 50px" maxlength="5" value="<%= GetMilitaryTime(SiteInfo.FridayOpenTime) %>" />
                                </td>
                                <td>
                                    <input type="text" name="txtSaturdayOpen" id="txtSaturdayOpen" style="width: 50px" maxlength="5" value="<%= GetMilitaryTime(SiteInfo.SaturdayOpenTime) %>" />
                                </td>
                                <td>
                                    <input type="text" name="txtSundayOpen" id="txtSundayOpen" style="width: 50px" maxlength="5" value="<%= GetMilitaryTime(SiteInfo.SundayOpenTime) %>" />
                                </td>
                            </tr>
                            <tr>
                                <th>Close</th>
                               <td>
                                    <input type="text" name="txtMondayClose" id="txtMondayClose" style="width: 50px" maxlength="5" value="<%= GetMilitaryTime(SiteInfo.MondayCloseTime) %>" />
                                </td>
                                <td>
                                    <input type="text" name="txtTuesdayClose" id="txtTuesdayClose" style="width: 50px" maxlength="5" value="<%= GetMilitaryTime(SiteInfo.TuesdayCloseTime) %>" />
                                </td>
                                <td>
                                    <input type="text" name="txtWednesdayClose" id="txtWednesdayClose" style="width: 50px" maxlength="5" value="<%= GetMilitaryTime(SiteInfo.WednesdayCloseTime) %>" />
                                </td>
                                <td>
                                    <input type="text" name="txtThursdayClose" id="txtThursdayClose" style="width: 50px" maxlength="5" value="<%= GetMilitaryTime(SiteInfo.ThursdayCloseTime) %>" />
                                </td>
                                <td>
                                    <input type="text" name="txtFridayClose" id="txtFridayClose" style="width: 50px" maxlength="5" value="<%= GetMilitaryTime(SiteInfo.FridayCloseTime) %>" />
                                </td>
                                <td>
                                    <input type="text" name="txtSaturdayClose" id="txtSaturdayClose" style="width: 50px" maxlength="5" value="<%= GetMilitaryTime(SiteInfo.SaturdayCloseTime) %>" />
                                </td>
                                <td>
                                    <input type="text" name="txtSundayClose" id="txtSundayClose" style="width: 50px" maxlength="5" value="<%= GetMilitaryTime(SiteInfo.SundayCloseTime) %>" />
                                </td>
                            </tr>
                        </table>

                        <%
                            if bSaved then
                        %>
                            <p class="center red">Settings saved successfully</p>
                        <%
                            end if
                        %>

                        <p id="valMessage" class="center red" style="display: none"></p>

                        <button type="submit" class="marginCenter" style="display: block" onclick="return validateForm()">Save Changes</button>
                    </td>
                </tr>
            </table>
        </form>
    </body>
</html>