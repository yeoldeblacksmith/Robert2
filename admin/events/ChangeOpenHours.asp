<!DOCTYPE html>
<!--#include file="../../classes/IncludeList.asp" -->
<% 
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForAdminOnly

    dim myDate, dateString
    set myDate = new AvailableDate
    myDate.Load(CDate(Request.QueryString("dt")))

    dateString = WeekdayName(Weekday(CDate(myDate.SelectedDate))) & ", " & _
                 MonthName(Month(myDate.SelectedDate)) & " " & _
                 Day(CDate(myDate.SelectedDate))
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title></title>

        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
            function closeWindow() {
                parent.$.fancybox.close();
            }

            function saveAvailableDate() {
                var myDay = '<%= myDate.SelectedDate %>';
                var openTime = $("select[name=openTime]").val();
                var closeTime = $("select[name=closeTime]").val();
                var url = '../../ajax/availability.asp?av=1&dt=' + myDay + '&st=' + openTime + '&et=' + closeTime
                $("input[value=Update]").attr("disabled", "disabled");

                $.ajax({
                    url:        url,
                    success:    function() {
                                    parent.loadCalendar(myDay);
                                    closeWindow();
                                },
                    error:      function () { alert('problems encountered'); }
                });
            }

            function removeAvailableDate() {
                var myDay = '<%= myDate.SelectedDate %>';
                var url = '../../ajax/availability.asp?av=0&dt=' + myDay;
                $("input[value=Remove]").attr("disabled", "disabled");

                $.ajax({
                    url:        url,
                    success:    function() {
                                    parent.loadCalendar(myDay);
                                    closeWindow();
                                },
                    error:      function () { alert('problems encountered'); }
                });
            }


<% 
    if myDate.Loaded then
%>
            $(document).ready(function () {
                $("select[name=openTime]").val("<%= myDate.StartTimeShortFormat%>");
                $("select[name=closeTime]").val("<%= myDate.EndTimeShortFormat%>");
            });
<%
    else
%>
            $(document).ready(function () {
                var dt = new Date('<%= myDate.SelectedDate %>');

                switch(dt.getDay()) {
                    case 0:
                        $("select[name=openTime]").val("<%= GetMilitaryTime(SiteInfo.SundayOpenTime) %>");
                        $("select[name=closeTime]").val("<%= GetMilitaryTime(SiteInfo.SundayCloseTime) %>");
                        break;
                    case 1:
                        $("select[name=openTime]").val("<%= GetMilitaryTime(SiteInfo.MondayOpenTime) %>");
                        $("select[name=closeTime]").val("<%= GetMilitaryTime(SiteInfo.MondayCloseTime) %>");
                        break;
                    case 2:
                        $("select[name=openTime]").val("<%= GetMilitaryTime(SiteInfo.TuesdayOpenTime) %>");
                        $("select[name=closeTime]").val("<%= GetMilitaryTime(SiteInfo.TuesdayCloseTime) %>");
                        break;
                    case 3:
                        $("select[name=openTime]").val("<%= GetMilitaryTime(SiteInfo.WednesdayOpenTime) %>");
                        $("select[name=closeTime]").val("<%= GetMilitaryTime(SiteInfo.WednesdayCloseTime) %>");
                        break;
                    case 4:
                        $("select[name=openTime]").val("<%= GetMilitaryTime(SiteInfo.ThursdayOpenTime) %>");
                        $("select[name=closeTime]").val("<%= GetMilitaryTime(SiteInfo.ThursdayCloseTime) %>");
                        break;
                    case 5:
                        $("select[name=openTime]").val("<%= GetMilitaryTime(SiteInfo.FridayOpenTime) %>");
                        $("select[name=closeTime]").val("<%= GetMilitaryTime(SiteInfo.FridayCloseTime) %>");
                        break;
                    case 6:
                        $("select[name=openTime]").val("<%= GetMilitaryTime(SiteInfo.SaturdayOpenTime) %>");
                        $("select[name=closeTime]").val("<%= GetMilitaryTime(SiteInfo.SaturdayCloseTime) %>");
                        break;
                }

                $("#btnUpdate").focus();
            });
<%
    end if
%>
        </script>
    </head>
    <body>
        <div style="text-align: center">
            <p>Select the hours of operation for <%= dateString  %></p>
            
            <table cellpadding="2" cellspacing="3" border="0" style="margin: 0 auto">
                <tr>
                    <td>Open:</td>
                    <td>
                        <select name="openTime">
                            <option value="00:00">12:00 AM</option>
                            <option value="00:30">12:30 AM</option>
                            <option value="01:00">1:00 AM</option>
                            <option value="01:30">1:30 AM</option>
                            <option value="02:00">2:00 AM</option>
                            <option value="02:30">2:30 AM</option>
                            <option value="03:00">3:00 AM</option>
                            <option value="03:30">3:30 AM</option>
                            <option value="04:00">4:00 AM</option>
                            <option value="04:30">4:30 AM</option>
                            <option value="05:00">5:00 AM</option>
                            <option value="05:30">5:30 AM</option>
                            <option value="06:00">6:00 AM</option>
                            <option value="06:30">6:30 AM</option>
                            <option value="07:00">7:00 AM</option>
                            <option value="07:30">7:30 AM</option>
                            <option value="08:00">8:00 AM</option>
                            <option value="08:30">8:30 AM</option>
                            <option value="09:00">9:00 AM</option>
                            <option value="09:30">9:30 AM</option>
                            <option value="10:00">10:00 AM</option>
                            <option value="10:30">10:30 AM</option>
                            <option value="11:00">11:00 AM</option>
                            <option value="11:30">11:30 AM</option>
                            <option value="12:00">12:00 PM</option>
                            <option value="12:30">12:30 PM</option>
                            <option value="13:00">1:00 PM</option>
                            <option value="13:30">1:30 PM</option>
                            <option value="14:00">2:00 PM</option>
                            <option value="14:30">2:30 PM</option>
                            <option value="15:00">3:00 PM</option>
                            <option value="15:30">3:30 PM</option>
                            <option value="16:00">4:00 PM</option>
                            <option value="16:30">4:30 PM</option>
                            <option value="17:00">5:00 PM</option>
                            <option value="17:30">5:30 PM</option>
                            <option value="18:00">6:00 PM</option>
                            <option value="18:30">6:30 PM</option>
                            <option value="19:00">7:00 PM</option>
                            <option value="19:30">7:30 PM</option>
                            <option value="20:00">8:00 PM</option>
                            <option value="20:30">8:30 PM</option>
                            <option value="21:00">9:00 PM</option>
                            <option value="21:30">9:30 PM</option>
                            <option value="22:00">10:00 PM</option>
                            <option value="22:30">10:30 PM</option>
                            <option value="23:00">11:00 PM</option>
                            <option value="23:30">11:30 PM</option>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td>Close:</td>
                    <td>
                        <select name="closeTime">
                            <option value="00:00">12:00 AM</option>
                            <option value="00:30">12:30 AM</option>
                            <option value="01:00">1:00 AM</option>
                            <option value="01:30">1:30 AM</option>
                            <option value="02:00">2:00 AM</option>
                            <option value="02:30">2:30 AM</option>
                            <option value="03:00">3:00 AM</option>
                            <option value="03:30">3:30 AM</option>
                            <option value="04:00">4:00 AM</option>
                            <option value="04:30">4:30 AM</option>
                            <option value="05:00">5:00 AM</option>
                            <option value="05:30">5:30 AM</option>
                            <option value="06:00">6:00 AM</option>
                            <option value="06:30">6:30 AM</option>
                            <option value="07:00">7:00 AM</option>
                            <option value="07:30">7:30 AM</option>
                            <option value="08:00">8:00 AM</option>
                            <option value="08:30">8:30 AM</option>
                            <option value="09:00">9:00 AM</option>
                            <option value="09:30">9:30 AM</option>
                            <option value="10:00">10:00 AM</option>
                            <option value="10:30">10:30 AM</option>
                            <option value="11:00">11:00 AM</option>
                            <option value="11:30">11:30 AM</option>
                            <option value="12:00">12:00 PM</option>
                            <option value="12:30">12:30 PM</option>
                            <option value="13:00">1:00 PM</option>
                            <option value="13:30">1:30 PM</option>
                            <option value="14:00">2:00 PM</option>
                            <option value="14:30">2:30 PM</option>
                            <option value="15:00">3:00 PM</option>
                            <option value="15:30">3:30 PM</option>
                            <option value="16:00">4:00 PM</option>
                            <option value="16:30">4:30 PM</option>
                            <option value="17:00">5:00 PM</option>
                            <option value="17:30">5:30 PM</option>
                            <option value="18:00">6:00 PM</option>
                            <option value="18:30">6:30 PM</option>
                            <option value="19:00">7:00 PM</option>
                            <option value="19:30">7:30 PM</option>
                            <option value="20:00">8:00 PM</option>
                            <option value="20:30">8:30 PM</option>
                            <option value="21:00">9:00 PM</option>
                            <option value="21:30">9:30 PM</option>
                            <option value="22:00">10:00 PM</option>
                            <option value="22:30">10:30 PM</option>
                            <option value="23:00">11:00 PM</option>
                            <option value="23:30">11:30 PM</option>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <input type="button" id="btnUpdate" value="Update" onclick="saveAvailableDate()" />
<% 
    if myDate.Loaded then
%>
                        &nbsp;
                        <input type="button" value="Remove" onclick="removeAvailableDate()" />
<%
    end if
%>
                        &nbsp;
                        <input type="button" value="Cancel" onclick="closeWindow()" />
                    </td>
                </tr>
            </table>
        </div>
    </body>
</html>
