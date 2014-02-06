<!DOCTYPE html>
<!--#include file="../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForAdminOnly

    dim bSaved
    bSaved = false

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        for each key in Request.Form
            settings(key) = request.Form(key)
            
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
    <title>Email Settings</title>
            <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/admin.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/navmenu.css" />

        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
            function validateEmailAddress() {
                var valid = true;
                var emailRegExp = /^([a-zA-Z0-9_\.\-\+])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
                var email = $("#EmailFromAddress").val();

                if (email.length == 0) {
                    valid = false;
                    $("#valEmailFromAddressReq").css("display", "block");
                    $("#valEmailFromAddressFormat").css("display", "none");
                } else if (!emailRegExp.test(email)) {
                    valid = false;
                    $("#valEmailFromAddressReq").css("display", "none");
                    $("#valEmailFromAddressFormat").css("display", "block");
                } else {
                    $("#valEmailFromAddressReq").css("display", "none");
                    $("#valEmailFromAddressFormat").css("display", "none");
                }

                return valid;
            }

            function validateForm() {
                return !!(validateEmailAddress() & validateServerAddress());
            }

            function validateServerAddress() {
                var valid = true;
                var ipRegEx = /^(([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$/;
                var address = $("#EmailServer").val();
                
                if (address.length == 0) {
                    valid = false;
                    $("#valEmailServerReq").css("display", "block");
                    $("#valEmailServerFormat").css("display", "none");
                } else if (!ipRegEx.test(address)) {
                    valid = false;
                    $("#valEmailServerReq").css("display", "none");
                    $("#valEmailServerFormat").css("display", "block");
                } else {
                    $("#valEmailServerReq").css("display", "none");
                    $("#valEmailServerFormat").css("display", "none");
                }

                return valid;
            }
        </script>
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Email Settings"

    navmenu.WriteNavigationSection NAVIGATION_NAME_SETTINGS 
%>
        <form method="post">
            <table class="adminTable" style="margin: 0 auto; width: 550px; height: auto;">
                <tr>
                    <td class="adminCell" style=" padding: 10px">
                        <table class="tableDefault marginCenter fullWidth" >
                            <tr class="normal">
                                <td>SMTP Server IP Address:</td>
                                <td>
                                    <input type="text" name="EmailServer" id="EmailServer" value="<%= settings(SETTING_EMAILSERVER) %>" class="fullWidth" onblur="validateServerAddress()" /><br />
                                    <span id="valEmailServerReq" class="hidden red">Required</span>
                                    <span id="valEmailServerFormat" class="hidden red">Invalid format</span>
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td>From Email Address:</td>
                                <td>
                                    <input type="text" name="EmailFromAddress" id="EmailFromAddress" value="<%= settings(SETTING_FROMADDRESS) %>" class="fullWidth" onblur="validateEmailAddress()" /><br />
                                    <span id="valEmailFromAddressReq" class="hidden red">Required</span>
                                    <span id="valEmailFromAddressFormat" class="hidden red">Invalid email address</span>
                                </td>
                            </tr>
                            <tr class="normal">
                                <td>From Email Name:</td>
                                <td>
                                    <input type="text" name="EmailFromName" id="EmailFromName" value="<%= settings(SETTING_FROMNAME) %>" class="fullWidth" /><br />
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td>Info Email Address:</td>
                                <td>
                                    <input type="text" name="InfoAddress" id="InfoAddress" value="<%= settings(SETTING_INFOADDRESS) %>" class="fullWidth" /><br />
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

                        <button type="submit" class="marginCenter" style="display: block" onclick="return validateForm()">Save Changes</button>
                    </td>
                </tr>
            </table>
        </form>
    </body>
</html>
