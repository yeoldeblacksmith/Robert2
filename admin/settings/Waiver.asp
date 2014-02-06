<!DOCTYPE html>
<!--#include file="../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForAdminOnly

    dim bSaved
    bSaved = false

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        for each key in Request.Form
            settings(key) = Request.Form(key)
        
            dim mySetting
            set mySetting = new SiteSetting
            mySetting.Key = key
            mySetting.Value = request.Form(key)
            mySetting.Save
        next

        bSaved = true
    end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Waiver Settings</title>
        <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/admin.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/navmenu.css" />

        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
            function validateInteger(FieldId, ValMessageId) {
                var valid = true;

                if (isNaN( $("#" + FieldId).val() )) {
                    valid = false;
                    $("#" + ValMessageId).css("display", "block");
                } else if ($("#" + FieldId).val() % 1 > 0) {
                    valid = false;
                    $("#" + ValMessageId).css("display", "block");
                } else {
                    $("#" + ValMessageId).css("display", "none");
                }

                return valid;
            }

            function validateRequired(FieldId, ValMessageId) {
                var valid = true;

                if ($("#" + FieldId).val().length == 0) {
                    valid = false;
                    $("#" + ValMessageId).css("display", "block");
                } else {
                    $("#" + ValMessageId).css("display", "none");
                }

                return valid;
            }

            function validateFieldBookingInterval() {
                return validateRequired("BookingInterval", "valBookingIntervalReq") && validateInteger("BookingInterval", "valBookingIntervalFormat");
            }

            function validateFieldMinAge() {
                return validateRequired("MinAge", "valMinAgeReq") && validateInteger("MinAge", "valMinAgeFormat");
            }

            function validateForm() {
                return !!(validateFieldMinAge() & validateFieldBookingInterval());
            }
        </script>
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Waiver Settings"

    navmenu.WriteNavigationSection NAVIGATION_NAME_SETTINGS 
%>
        <form method="post">
            <table class="adminTable" style="margin: 0 auto; width: 550px; height: auto;">
                <tr>
                    <td class="adminCell" style=" padding: 10px">
                        <table class="tableDefault marginCenter border">
                            <tr class="normal">
                                <td style="width: 250px">Minimum Age:</td>
                                <td>
                                    <input type="text" name="<%= SETTING_WAIVER_MINAGE %>" id="MinAge" value="<%= Settings(SETTING_WAIVER_MINAGE) %>" /><br />
                                    <span id="valMinAgeReq" class="red" style="display: none">Required</span>
                                    <span id="valMinAgeFormat" class="red" style="display: none">Invalid format</span>
                                </td>
                            </tr>
                            <tr class="normal">
                                <td colspan="2">
                                    Note: This is the minimum required age to fill out a waiver. If no minimum is required, change the value to zero.
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td>Booking Interval (in minutes):</td>
                                <td>
                                    <input type="text" name="<%= SETTING_BOOKING_INTERVAL %>" id="BookingInterval" value="<%= Settings(SETTING_BOOKING_INTERVAL) %>" /><br />
                                    <span id="valBookingIntervalReq" class="red" style="display: none">Required</span>
                                    <span id="valBookingIntervalFormat" class="red" style="display: none">Invalid format</span>
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td colspan="2">
                                    Note: This changes the amount of time between time slots in registration and waivers (i.e. 30 minutes = 10:00 - 10:30).
                                </td>
                            </tr>

                            <tr class="normal">
                                <td colspan="2">Waiver Legal Disclaimer:</td>
                            </tr>
                            <tr class="normal">
                                <td colspan="2">
                                    <% 
                                        Dim bodyEditor
	                                    Set bodyEditor = New CuteEditor
                                        bodyEditor.FilesPath = "../../../../cute/CuteEditor_Files"
	                                    bodyEditor.HelpUrl = "help.asp"
                                        bodyEditor.ImageGalleryPath = "../../content/images/cuteupload"
                                        bodyEditor.ConfigurationPath = "../../../../cute/CuteEditor_Files/Configuration/autoconfigure/simple.config"
                                        bodyEditor.Id = SETTING_WAIVER_LEGALESE
                                        bodyEditor.Text = settings(SETTING_WAIVER_LEGALESE)
                                        bodyEditor.Draw()
                                    %>
                                </td>
                            </tr>

                            <tr class="alternate">
                                <td colspan="2">Blurb:</td>
                            </tr>
                            <tr class="normal">
                                <td colspan="2">
                                    <% 
                                        Dim bodyEditor1
	                                    Set bodyEditor1 = New CuteEditor
                                        bodyEditor1.FilesPath = "../../../../cute/CuteEditor_Files"
	                                    bodyEditor1.HelpUrl = "help.asp"
                                        bodyEditor1.ImageGalleryPath = "../../content/images/cuteupload"
                                        bodyEditor1.ConfigurationPath = "../../../../cute/CuteEditor_Files/Configuration/autoconfigure/simple.config"
                                        bodyEditor1.Id = SETTING_BLURB_WAIVER
                                        bodyEditor1.Text = settings(SETTING_BLURB_WAIVER)
                                        bodyEditor1.Draw()
                                    %>
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

                        <button type="submit" class="marginCenter" style="display: block" onclick="return validateForm();">Save Changes</button>
                    </td>
                </tr>
            </table>
        </form>
    </body>
</html>
