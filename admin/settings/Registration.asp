<!DOCTYPE html>
<!--#include file="../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForAdminOnly

    dim bSaved
    bSaved = false

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        dim moduleKeyFound : moduleKeyFound = false

        for each key in Request.Form
            ' if the key is sent, it has to be true
            if key = SETTING_MODULE_REGISTRATION then
                settings(key) = "true"
                moduleKeyFound = true
            else
                settings(key) = Request.Form(key)
            end if

            if key = SETTING_BOOKING_INTERVAL then
                settings(key) = abs(Request.Form(key))
            end if

            dim mySetting
            set mySetting = new SiteSetting
            mySetting.Key = key
            mySetting.Value = settings(key)
            mySetting.Save
        next

        if moduleKeyFound = false then
            settings(SETTING_MODULE_REGISTRATION) = "false"
            'dim mySetting
            set mySetting = new SiteSetting
            mySetting.Key = SETTING_MODULE_REGISTRATION
            mySetting.Value = "false"
            mySetting.Save
        end if

        bSaved = true
    end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Registration Settings</title>
        <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/admin.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/navmenu.css" />

        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
            function validateDouble(FieldId, ValMessageId) {
                var valid = true;

                if (isNaN($("#" + FieldId).val())) {
                    valid = false;
                    $("#" + ValMessageId).css("display", "block");
                } else {
                    $("#" + ValMessageId).css("display", "none");
                }

                return valid;
            }

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

            function validateInterval() {
                var valid = true;

                if (parseInt($("#BookingInterval").val()) % 15 == 0) {
                    $("#valBookingIntervalCustom").css("display", "none");
                } else {
                    valid = false;
                    $("#valBookingIntervalCustom").css("display", "block");
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

            function validateFieldDaysOut() {
                return validateRequired("DaysOut", "valDaysOutReq") && validateInteger("DaysOut", "valDaysOutFormat");
            }

            function validateFieldDisplayMonths() {
                return validateRequired("DisplayMonths", "valDisplayMonthsReq") && validateInteger("DisplayMonths", "valDisplayMonthsFormat");
            }

            function validateFieldBookingInterval() {
                return validateRequired("BookingInterval", "valBookingIntervalReq") &&
                       validateInteger("BookingInterval", "valBookingIntervalFormat") &&
                       validateInterval();
            }

            function validateFieldMaxPlayers() {
                return validateRequired("MaxPlayers", "valMaxPlayersReq") && validateInteger("MaxPlayers", "valMaxPlayersFormat");
            }

            function validateFieldMinPlayers() {
                return validateRequired("MinPlayers", "valMinPlayersReq") && validateInteger("MinPlayers", "valMinPlayersFormat");
            }

            function validateFieldPaymentAmount() {
                var valid = !!(validateRequired("PaymentAmount", "valPaymentAmountReq") && validateDouble("PaymentAmount", "valPaymentAmountFormat"));

                if (valid) {
                    if ($("#PaymentType").val() != "0" && parseFloat($("#PaymentAmount").val()) == 0) {
                        valid = false;
                        $("#valPaymentAmountReq").css("display", "block");
                    } else {
                        $("#valPaymentAmountReq").css("display", "none");
                    }
                }

                return valid;
            }

            function validateFieldPlayerLimit() {
                return validateRequired('PlayerLimit', 'valPlayerLimitReq') && validateInterval('PlayerLimit', 'valPlayerLimitFormat');
            }

            function validateForm() {
                return !!(validateFieldMaxPlayers() & validateFieldDaysOut() & validateFieldBookingInterval() &
                          validateFieldPlayerLimit() & validateFieldMinPlayers() & validateFieldPaymentAmount() &
                          validateFieldDisplayMonths());
            }
        </script>
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Registration Settings"

    navmenu.WriteNavigationSection NAVIGATION_NAME_SETTINGS 
%>
        <form method="post">
            <table class="adminTable" style="margin: 0 auto; width: 550px; height: auto;">
                <tr>
                    <td class="adminCell" style=" padding: 10px">
                        <table class="tableDefault marginCenter border">
                            <tr class="normal">
                                <td>Module Enabled</td>
                                <td>
                                    <input type="checkbox" name="<%= SETTING_MODULE_REGISTRATION %>" id="chkEnabled" <%= iif(cbool(Settings(SETTING_MODULE_REGISTRATION)), "checked=""checked""", "") %> />
                                </td>
                            </tr>
                            <tr class="normal">
                                <td colspan="2">
                                    Note: Changing this setting will disable the registration module. Customers will no longer be able scheduled events. Hooks from the 
                                    waiver module that tie back to the registration module will also be removed.
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td>Months to Display</td>
                                <td>
                                    <input type="text" name="<%= SETTING_REGISTRATION_DISPLAYMONTHS %>" id="DisplayMonths" value="<%= Settings(SETTING_REGISTRATION_DISPLAYMONTHS) %>" /><br />
                                    <span id="valDisplayMonthsReq" class="red" style="display: none">Required</span>
                                    <span id="valDisplayMonthsFormat" class="red" style="display: none">Invalid number</span>
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td colspan="2">
                                    Note: This setting ontrols how many months are displayed on the calendar for customers when booking their registration.
                                </td>
                            </tr>
                            <tr class="normal">
                                <td>Max Allowed Per Time Slot (soft setting):</td>
                                <td>
                                    <input type="text" name="<%= SETTING_REGISTRATION_MAXPLAYERS%>" id="MaxPlayers" value="<%= Settings(SETTING_REGISTRATION_MAXPLAYERS)%>" /><br />
                                    <span id="valMaxPlayersReq" class="red" style="display: none">Required</span>
                                    <span id="valMaxPlayersFormat" class="red" style="display: none">Invalid format</span>
                                </td>
                            </tr>
                            <tr class="normal">
                                <td colspan="2">
                                    Note: Max Allower Per Time Slot determines how many people can register for an event during the same time slot. If the limit
                                    has been exceeded, the time slot will not show up as an option when users are registering.
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td>Total Max Players at Facility:</td>
                                <td>
                                    <input type="text" name="<%= SETTING_REGISTRATION_PLAYERLIMIT%>" id="PlayerLimit" value="<%= Settings(SETTING_REGISTRATION_PLAYERLIMIT)%>" /><br />
                                    <span id="valPlayerLimitReq" class="red" style="display: none">Required</span>
                                    <span id="valPlayerLimitFormat" class="red" style="display: none">Invalid format</span>
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td colspan="2">
                                    Note: Combined with your party duration below, you can set the maximum number of players your facility can handle. Example: With only 100 sets of equipment, 
                                    you could set this for 100, then if 40 book in at 10:00 and 40 book in at 10:30, and the party duration is for 2 hours, for the next hour, only 20 more 
                                    could book in. After noon, 40 more would be available as the 10 AM group would be done.
                                </td>
                            </tr>
                            <tr class="normal">
                                <td>Min Allowed For Reservation:</td>
                                <td>
                                    <input type="text" name="<%= SETTING_REGISTRATION_MINPLAYERS%>" id="MinPlayers" value="<%= Settings(SETTING_REGISTRATION_MINPLAYERS)%>" /><br />
                                    <span id="valMinPlayersReq" class="red" style="display: none">Required</span>
                                    <span id="valMinPlayersFormat" class="red" style="display: none">Invalid format</span>
                                </td>
                            </tr>
                            <tr class="normal">
                                <td colspan="2">
                                    Note: Min Allowed For Reservation will set the minimum players required to make a reservation. Setting the value to 1 or 0 will eliminate the minimum.
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td>Reminder Days Out:</td>
                                <td>
                                    <input type="text" name="<%= SETTING_REGISTRATION_REMINDERDAYS%>" id="DaysOut" value="<%= Settings(SETTING_REGISTRATION_REMINDERDAYS)%>" /><br />
                                    <span id="valDaysOutReq" class="red" style="display: none">Required</span>
                                    <span id="valDaysOutFormat" class="red" style="display: none">Invalid format</span>
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td colspan="2">
                                    Note: Reminder Days Out tells the system how far out to pull for events before sending the contacts a reminder email for their
                                    reservation.
                                </td>
                            </tr>
                            <tr class="normal">
                                <td>Booking Interval (in minutes):</td>
                                <td>
                                    <input type="text" name="<%= SETTING_BOOKING_INTERVAL %>" id="BookingInterval" value="<%= Settings(SETTING_BOOKING_INTERVAL) %>" /><br />
                                    <span id="valBookingIntervalReq" class="red" style="display: none">Required</span>
                                    <span id="valBookingIntervalFormat" class="red" style="display: none">Invalid format</span>
                                    <span id="valBookingIntervalCustom" class="red" style="display: none">Please use 15 minute increments</span>
                                </td>
                            </tr>
                            <tr class="normal">
                                <td colspan="2">
                                    Note: This changes the amount of time between time slots in registration and waivers (i.e. 30 minutes = 10:00 - 10:30).
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td>Expected Party Duration (in minutes):</td>
                                <td>
                                    <input type="text" name="<%= SETTING_REGISTRATION_PARTYDURATION %>" id="PartyDuration" value="<%= Settings(SETTING_REGISTRATION_PARTYDURATION) %>" /><br />
                                    <span id="valPartyDurationReq" class="red" style="display: none">Required</span>
                                    <span id="valPartyDurationFormat" class="red" style="display: none">Invalid format</span>
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td colspan="2">
                                    Note: How long does a typical party stay? This will be taken into consideration when allowing users to register for parties.
                                </td>
                            </tr>
                            <tr class="normal">
                                <td>Deposit Amount:</td>
                                <td>
                                    <input type="text" name="<%= SETTING_REGISTRATION_PAYMENTAMOUNT %>" id="PaymentAmount" value="<%= Settings(SETTING_REGISTRATION_PAYMENTAMOUNT) %>" /><br />
                                    <span id="valPaymentAmountReq" class="red" style="display: none">Required</span>
                                    <span id="valPaymentAmountFormat" class="red" style="display: none">Not a number (i.e. 100.00)</span>
                                </td>
                            </tr>
                            <tr class="normal">
                                <td colspan="2">
                                    Note: Amount charged depends on the payment type selected below. If payment type is set to none, there will not be a calculation performed. If the
                                    payment type is set to by event, then a flat rate of the deposit amount will be assessed. If by player is selected, then the deposit amount will be
                                    multiplied by the number of players.
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td>Payment Type</td>
                                <td>
                                    <% BuildPaymentTypeList Settings(SETTING_REGISTRATION_PAYMENTTYPE) %>
                                </td>
                            </tr>
                            <tr class="normal">
                                <td colspan="2">&nbsp;</td>
                            </tr>
                            <tr class="alternate">
                                <td colspan="2">Calendar Blurb:</td>
                            </tr>
                            <tr class="normal">
                                <td colspan="2">
                                    <% 
                                        Dim calEditor
	                                    Set calEditor = New CuteEditor
                                        calEditor.FilesPath = "../../../../cute/CuteEditor_Files"
	                                    calEditor.HelpUrl = "help.asp"
                                        calEditor.ImageGalleryPath = "../../content/images/cuteupload"
                                        calEditor.ConfigurationPath = "../../../../cute/CuteEditor_Files/Configuration/autoconfigure/simple.config"
                                        calEditor.Id = SETTING_BLURB_CALENDAR
                                        calEditor.Text = settings(SETTING_BLURB_CALENDAR)
                                        calEditor.Draw()
                                    %>
                                </td>
                            </tr>

                            <tr class="alternate">
                                <td colspan="2">Registration Blurb:</td>
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
                                        bodyEditor.Id = SETTING_BLURB_REGISTRATION
                                        bodyEditor.Text = settings(SETTING_BLURB_REGISTRATION)
                                        bodyEditor.Draw()
                                    %>
                                </td>
                            </tr>

                            <tr class="alternate">
                                <td colspan="2">Confirmation Blurb:</td>
                            </tr>
                            <tr class="normal">
                                <td colspan="2">
                                    <% 
                                        Dim confirmEditor
	                                    Set confirmEditor = New CuteEditor
                                        confirmEditor.FilesPath = "../../../../cute/CuteEditor_Files"
	                                    confirmEditor.HelpUrl = "help.asp"
                                        confirmEditor.ImageGalleryPath = "../../content/images/cuteupload"
                                        confirmEditor.ConfigurationPath = "../../../../cute/CuteEditor_Files/Configuration/autoconfigure/simple.config"
                                        confirmEditor.Id = SETTING_BLURB_CONFIRMATION
                                        confirmEditor.Text = settings(SETTING_BLURB_CONFIRMATION)
                                        confirmEditor.Draw()
                                    %>
                                </td>
                            </tr>
                            
                            <tr class="alternate">
                                <td colspan="2">Terms and Conditions:</td>
                            </tr>
                            <tr class="normal">
                                <td colspan="2">
                                    <% 
                                        Dim termsEditor
	                                    Set termsEditor = New CuteEditor
                                        termsEditor.FilesPath = "../../../../cute/CuteEditor_Files"
	                                    termsEditor.HelpUrl = "help.asp"
                                        termsEditor.ImageGalleryPath = "../../content/images/cuteupload"
                                        termsEditor.ConfigurationPath = "../../../../cute/CuteEditor_Files/Configuration/autoconfigure/simple.config"
                                        termsEditor.Id = SETTING_BLURB_REGISTRATION_TERMS
                                        termsEditor.Text = settings(SETTING_BLURB_REGISTRATION_TERMS)
                                        termsEditor.Draw()
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
