<!DOCTYPE html>
<!--#include file="../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForAdminOnly

    dim bSaved
    bSaved = false

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        dim cancelKeyFound : cancelKeyFound = false
        dim editKeyFound : editKeyFound = false

        for each key in Request.Form
            ' if the key is sent, it has to be true
            if key = SETTING_REGISTRATION_ALLOWCANCELLATION then
                settings(key) = "true"
                cancelKeyFound = true
            elseif key = SETTING_REGISTRATION_ALLOWEDIT then
                settings(key) = "true"
                editKeyFound = true
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

        if cancelKeyFound = false then
            settings(SETTING_REGISTRATION_ALLOWCANCELLATION) = "false"
            set mySetting = new SiteSetting
            mySetting.Key = SETTING_REGISTRATION_ALLOWCANCELLATION
            mySetting.Value = "false"
            mySetting.Save
        end if

        if editKeyFound = false then
            settings(SETTING_REGISTRATION_ALLOWEDIT) = "false"
            set mySetting = new SiteSetting
            mySetting.Key = SETTING_REGISTRATION_ALLOWEDIT
            mySetting.Value = "false"
            mySetting.Save
        end if

        bSaved = true
    end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Registration Cancel/Modify Settings</title>
        <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/admin.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/navmenu.css" />

        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
            function validateInteger(FieldId, ValMessageId) {
                var valid = true;

                if (isNaN($("#" + FieldId).val())) {
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

            function validateFieldCancelDays(){ 
                return !!(validateRequired('txtCancelDays','valCancelDaysReq') && validateInteger('txtCancelDays','valCancelDaysFormat'));
            }

            function validateFieldEditDays() {
                return !!(validateRequired('txtEditDays','valEditDaysReq') && validateInteger('txtEditDays','valEditDaysFormat'));
            }

            function validateForm() {
                var validCancel = true;

                if ($("#chkCancel").is(":checked")) {
                    validCancel = !!(validateFieldCancelDays());
                }

                var validEdit = true;

                if ($("#chkEdit").is(":checked")) {
                    validEdit = !!(validateFieldEditDays());
                }

                return !!(validCancel & validEdit);
            }
        </script>
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Registration Cancel/Modify Settings"

    navmenu.WriteNavigationSection NAVIGATION_NAME_SETTINGS 
%>
        <form method="post">
            <table class="adminTable" style="margin: 0 auto; width: 550px; height: auto;">
                <tr>
                    <td class="adminCell" style=" padding: 10px">
                        <table class="tableDefault marginCenter border">
                            <tr class="normal">
                                <td>Allow Cancellation</td>
                                <td>
                                    <input type="checkbox" id="chkCancel" name="<%= SETTING_REGISTRATION_ALLOWCANCELLATION %>" <%= iif(cbool(Settings(SETTING_REGISTRATION_ALLOWCANCELLATION)), "checked=""checked""", "") %> />
                                </td>
                            </tr>
                            <tr class="normal">
                                <td colspan="2">
                                    Note: This setting will allow your guests to cancel their reservation. Using the setting below, you can limit how far out they are allowed to
                                    cancel their party.
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td>Allow Modification</td>
                                <td>
                                    <input type="checkbox" id="chkEdit" name="<%= SETTING_REGISTRATION_ALLOWEDIT %>" <%= iif(cbool(Settings(SETTING_REGISTRATION_ALLOWEDIT)), "checked=""checked""", "") %> />
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td colspan="2">
                                    Note: This setting will allow your guests to modify their reservation. Using the setting below, you can limit how far out they are allowed 
                                    to edit their party.
                                </td>
                            </tr>
                            <tr class="normal">
                                <td>Days to Cancel</td>
                                <td>
                                    <input id="txtCancelDays" name="<%=SETTING_REGISTRATION_MAXDAYS_CANCEL %>" value="<%= Settings(SETTING_REGISTRATION_MAXDAYS_CANCEL)%>" />
                                    <span id="valCancelDaysReq" style="display: none; color: red">Required</span>
                                    <span id="valCancelDaysFormat" style="display: none; color: red">Invalid number</span>
                                </td>
                            </tr>
                            <tr class="normal">
                                <td colspan="2">
                                    Note: This setting dictates how many days prior to the selected date to allow the guest to cancel their reservation.
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td>Days to Modify</td>
                                <td>
                                    <input id="txtEditDays" name="<%=SETTING_REGISTRATION_MAXDAYS_EDIT%>" value="<%= Settings(SETTING_REGISTRATION_MAXDAYS_EDIT)%>" />
                                    <span id="valEditDaysReq" style="display: none; color: red">Required</span>
                                    <span id="valEditDaysFormat" style="display: none; color: red">Invalid number</span>
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td colspan="2">
                                    Note: This setting dictates how many days prior to the selected date to allow the guest to modify their reservation.
                                </td>
                            </tr>
                            <tr class="normal">
                                <td colspan="2">Denied Cancel/Modify Blurb</td>
                            </tr>
                            <tr class="normal">
                                <td colspan="2">
                                    <% 
                                        Dim editor
	                                    Set editor = New CuteEditor
                                        editor.FilesPath = "../../../../cute/CuteEditor_Files"
	                                    editor.HelpUrl = "help.asp"
                                        editor.ImageGalleryPath = "../../content/images/cuteupload"
                                        editor.ConfigurationPath = "../../../../cute/CuteEditor_Files/Configuration/autoconfigure/simple.config"
                                        editor.Id = SETTING_BLURB_REGISTRATION_CANCELDENIED
                                        editor.Text = settings(SETTING_BLURB_REGISTRATION_CANCELDENIED)
                                        editor.Draw()
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