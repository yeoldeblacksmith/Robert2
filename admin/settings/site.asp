<!DOCTYPE html>
<!--#include file="../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForAdminOnly

    dim bSaved
    bSaved = false

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        with SiteInfo
            .Name = Request.Form("txtName")
            .HomeUrl = Request.Form("txtHomeUrl")
            .VantoraUrl = Request.Form("txtVantoraUrl")
            .Address = request.Form("txtAddress")
            .City = request.Form("txtCity")
            .State = request.Form("ddlState")
            .ZipCode = request.Form("txtZipCode")
            .PhoneNumber = request.Form("txtPhoneNumber")
            .Country = request.Form("ddlCountry")

            .MondayOpenTime = GetMilitaryTime(.MondayOpenTime)
            .MondayCloseTime = GetMilitaryTime(.MondayCloseTime)
            .TuesdayOpenTime = GetMilitaryTime(.TuesdayOpenTime)
            .TuesdayCloseTime = GetMilitaryTime(.TuesdayCloseTime)
            .WednesdayOpenTime = GetMilitaryTime(.WednesdayOpenTime)
            .WednesdayCloseTime = GetMilitaryTime(.WednesdayCloseTime)
            .ThursdayOpenTime = GetMilitaryTime(.ThursdayOpenTime)
            .ThursdayCloseTime = GetMilitaryTime(.ThursdayCloseTime)
            .FridayOpenTime = GetMilitaryTime(.FridayOpenTime) 
            .FridayCloseTime = GetMilitaryTime(.FridayCloseTime) 
            .SaturdayOpenTime = GetMilitaryTime(.SaturdayOpenTime) 
            .SaturdayCloseTime = GetMilitaryTime(.SaturdayCloseTime)
            .SundayOpenTime = GetMilitaryTime(.SundayOpenTime)
            .SundayCloseTime = GetMilitaryTime(.SundayCloseTime)
            
            .Save
        end with

        dim mySetting
        set mySetting = new SiteSetting
        mySetting.Key = SETTING_PAYMENT_PAYPALADDRESS
        mySetting.Value = request.Form(SETTING_PAYMENT_PAYPALADDRESS)
        mySetting.Save

        ' reload the settings
        set Settings = GetSettingsDictionary()

        bSaved = true
    end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Site Settings</title>
        <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/admin.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/navmenu.css" />

        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">

        </script>
    </head>
    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Site Settings"

    navmenu.WriteNavigationSection NAVIGATION_NAME_SETTINGS 
%>
        <form method="post">
            <table class="adminTable" style="margin: 0 auto; width: 550px; height: auto;">
                <tr>
                    <td class="adminCell" style=" padding: 10px">
                        <h3>Site Info</h3>
                        <table class="tableDefault marginCenter border">
                            <tr class="normal border">
                                <td>Site GUID:</td>
                                <td><%= SiteInfo.SiteGuid %></td>
                            </tr>
                            <tr class="alternate border">
                                <td>Name:</td>
                                <td>
                                    <input type="text" name="txtName" id="txtName" value="<%= SiteInfo.Name %>" maxlength="200" />
                                </td>
                            </tr>
                            <tr class="normal border">
                                <td>Home URL:</td>
                                <td>
                                    <input type="text" name="txtHomeUrl" id="txtHomeUrl" value="<%= SiteInfo.HomeUrl %>" maxlength="256" style="width: 98%" />
                                </td>
                            </tr>
                            <tr class="alternate border">
                                <td>Vantora URL:</td>
                                <td>
                                    <%=SiteInfo.VantoraUrl %>
                                    <input type="hidden" name="txtVantoraUrl" id="txtVantoraUrl" value="<%= SiteInfo.VantoraUrl %>" />
                                </td>
                            </tr>
                            <tr class="normal border">
                                <td>Country:</td>
                                <td>
                                    <% BuildCountryDropDownList "ddlCountry", SiteInfo.Country %>
                                </td>
                            </tr>
                            <tr class="alternate border">
                                <td>Address:</td>
                                <td>
                                    <input type="text" name="txtAddress" id="txtAddress" value="<%= SiteInfo.Address %>" maxlength="256" style="width: 98%" />
                                </td>
                            </tr>
                            <tr class="normal border">
                                <td>City:</td>
                                <td>
                                    <input type="text" name="txtCity" id="txtCity" value="<%= SiteInfo.City %>" maxlength="50" style="width: 98%" />
                                </td>
                            </tr>
                            <tr class="alternate border">
                                <td><% 
                                        select case lcase(SiteInfo.Country)
                                            case "ca"
                                                response.Write "Province:"
                                            case "us"
                                                response.Write "State:"
                                        end select
                                    %></td>
                                <td>
                                    <% BuildStateDropDownList "ddlState", SiteInfo.State, SiteInfo.Country %>
                                </td>
                            </tr>
                            <tr class="normal border">
                                <td><% 
                                        select case lcase(SiteInfo.Country)
                                            case "ca"
                                                response.Write "Postal Code:"
                                            case "us"
                                                response.Write "Zip Code:"
                                        end select
                                    %></td>
                                <td>
                                    <input type="text" name="txtZipCode" id="txtZipCode" value="<%= SiteInfo.ZipCode %>" maxlength="20" />
                                </td>
                            </tr>
                            <tr class="alternate border">
                                <td>Phone Number:</td>
                                <td>
                                    <input type="text" name="txtPhoneNumber" id="txtPhoneNumber" value="<%= SiteInfo.PhoneNumber %>" maxlength ="20" />
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td class="adminCell" style="padding: 10px">
                        <h3>Payment Info:</h3>
                        <table class="tableDefault marginCenter border">
                            <tr>
                                <td>PayPal Address:</td>
                                <td>
                                    <input type="text" name="<%= SETTING_PAYMENT_PAYPALADDRESS%>" id="<%= SETTING_PAYMENT_PAYPALADDRESS %>" value="<%= Settings(SETTING_PAYMENT_PAYPALADDRESS) %>" maxlength="128" style="width: 98%" />
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    Note: Include the PayPal email address for your site to allow your customers to pay online
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td class="adminCell">
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
<%
    sub BuildCountryDropDownList(FieldName, SelectedCountry)
        with response
            .Write "<select name=""" & FieldName & """ id=""" & FieldName & """>" & vbCrLf
            .Write "<option value=""US""" & iif(UCase(SelectedCountry) = "US", " selected=""selected""", "") & ">United States</option>" & vbCrLf
            .Write "<option value=""CA""" & iif(UCase(SelectedCountry) = "CA", " selected=""selected""", "") & ">Canada</option>" & vbCrLf
            .Write "</select>" & vbCrLf
        end with
    end sub
%>