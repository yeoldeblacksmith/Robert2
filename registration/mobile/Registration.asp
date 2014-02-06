<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="../common/EventIncludeList.asp" -->
<%
    dim myDate
    set myDate = new AvailableDate
    myDate.Load(CDate(Request.Form("selectedDate")))

%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Event Registration</title>
        <meta name="viewport" content="width=device-width, initial-scale=1">
		<meta name="apple-mobile-web-app-capable" content="yes" />
	    
        <link rel="stylesheet" href="../content/css/jquery.mobile-1.1.0.css" />
		
        <style type="text/css">
			.center {
				text-align: center;
			}
			.center * {
				margin: 0 auto;
			}
		</style>

        <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
		<script type="text/javascript" src="../content/js/jquery.mobile-1.1.0.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.cookie.js"></script>
		<script type="text/javascript" src="../content/js/common.js"></script>
        <script type="text/javascript">
            String.prototype.trim = function () {
                return this.replace(/^\s*/, "").replace(/\s*$/, "");
            }

            function handleMaxLength(length) {
                var commentField = $("textarea[name=UserComments]");

                if (commentField.val().length <= commentField.attr("maxlength")) {
                    var remains = commentField.attr("maxlength") - commentField.val().length;
                    $("span[name=charcounter]").html("(" + remains + " characters available)");
                    return true;
                } else {
                    return false;
                }
            }

            function saveRegistration() {
                $("#btnSave").attr("disabled", "disabled");
                if (validateForm()) {
                    var url = '../common/eventajax.asp?act=0&dt=' + escape($("span[name=Date]").html()) +
                                                    '&st=' + escape($("select[name=Time]").val()) +
                                                    '&np=' + escape($("input[name=NumberOfGuests]").val()) +
                                                    '&ap=' + escape($("input[name=AgeOfGuests]").val().trim()) +
                                                    '&em=' + escape($("input[name=Email]").val().trim()) +
                                                    '&ph=' + escape($("input[name=Phone]").val().trim()) +
                                                    '&nm=' + escape($("input[name=Name]").val().trim()) +
                                                    '&ev=' + escape($("select[name=EventTypeId]").val()) +
                                                    '&pn=' + escape($("input[name=PartyName]").val().trim()) +
                                                    '&uc=' + escape($("textarea[name=UserComments]").val().trim());

                    $.ajax({
                        url: url,
                        success: function () { window.location = "Submitted.htm"; },
                        error: function () { alert("problems encountered"); }
                    });
                } else {
                    $("#btnSave").removeAttr("disabled");
                }
            }

            function validateForm() {
                var valid = true;

                // validate contact email address
                var email = $("input[name=Email]").val().toLowerCase();
                var confirm = $("input[name=ConfirmEmail]").val().toLowerCase();
                var emailRegExp = /^([a-zA-Z0-9_\.\-\+])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
                var phoneRegExp = /^\(?(\d{3})\)?[- ]?(\d{3})[- ]?(\d{4})$/;

                if (email.trim().length > 0) {
                    if (emailRegExp.test(email)) {
                        $("span[name=ValEmail]").css("display", "none");
                        $("span[name=ValEmailRequired]").css("display", "none");

                        if (email != confirm) {
                            $("span[name=ValEmailMatch]").css("display", "inline");
                            valid = false;
                        } else {
                            $("span[name=ValEmailMatch]").css("display", "none");
                        } // confirm email matching
                    } else {
                        $("span[name=ValEmail]").css("display", "inline");
                        valid = false;
                    } // regExp test
                } else {
                    $("span[name=ValEmailRequired]").css("display", "inline");
                    valid = false;
                } // length test

                var phone = $("input[name=Phone]").val();

                if (phone.trim().length == 0) {
                    $("span[name=ValPhoneRequired]").css("display", "inline");
                    valid = false;
                } else {
                    $("span[name=ValPhoneRequired]").css("display", "none");

                    if (phoneRegExp.test(phone)) {
                        $("span[name=ValPhoneFormat]").css("display", "none");
                    } else {
                        $("span[name=ValPhoneFormat]").css("display", "inline");
                    }
                }

                // validate contact name
                var name = $("input[name=Name]").val();

                if (name.trim().length == 0) {
                    $("span[name=ValName]").css("display", "inline");
                    valid = false;
                } else {
                    $("span[name=ValName]").css("display", "none");
                }

                // validate number of guests
                var guestCnt = $("input[name=NumberOfGuests]").val();

                if (guestCnt.trim().length == 0) {
                    $("span[name=ValNumberOfGuestsRequired]").css("display", "inline");
                    $("span[name=ValNumberOfGuests]").css("display", "none");
                    valid = false;
                } else {
                    $("span[name=ValNumberOfGuestsRequired]").css("display", "none");
                    if (!isInt(guestCnt)) {
                        $("span[name=ValNumberOfGuests]").css("display", "inline");
                        valid = false;
                    } else {
                        if (guestCnt <= 0) {
                            $("span[name=ValNumberOfGuests]").css("display", "inline");
                            valid = false;
                        } else {
                            $("span[name=ValNumberOfGuests]").css("display", "none");
                        }
                    }
                }

                return valid;
            }
        </script>

    </head>
    
    <body>
        <!-- Start of first page: #home -->
		<div data-role="page" id="home" data-theme="b">
            <!-- header -->
            <div data-role="header">
				<a href ="http://www.gatsplat.com/mobile/index.asp" rel="external" data-role="button" data-icon="home">Home</a><h1>Gatsplat</h1>
			</div>
			
            <!-- logo -->
            <div class="center" style="background: #ffffff; box-shadow: 0px 7px 5px #2d2d2d;">
                <img style="max-height: 200px; " src="../content/images/mobile/logo.png" alt="Gatsplat logo"/>
			</div>
			
            <!-- content -->
            <div data-role="content">
                <p style="text-align: center">Please Completely Fill Out The Registration Form</p>
    
                <table cellpadding="2" cellspacing="3" border="0" style="margin: 0 auto;">
                    <tr>
                        <td>Date:</td>
                        <td>
                            <span name="Date"><%= Request.Form("selectedDate") %></span>
                        </td>
                        <td/>
                    </tr>
                     <tr>
                        <td>Your Name:</td>
                        <td>
                            <input type="text" name="Name" maxlength="100" /><br />
                            <span name="ValName" style="display: none; color: Red">Name is required</span>
                        </td>
                        <td style="color: Red">*</td>
                    </tr>
                    <tr>
                        <td>What Time Do you Want to Start?:</td>
                        <td>
                            <select name="Time">
                                <% BuildTimeOptions myDate.StartTime, myDate.EndTime  %>
                            </select>
                        </td>
                        <td/>
                    </tr>
                    <tr>
                        <td>Approximate Number Of Guests:</td>
                        <td>
                            <!--
                            <select name="NumberOfGuests">
                                <option value=""></option>
                                <option value="2">2</option>
                                <option value="3">3</option>
                                <option value="4">4</option>
                                <option value="5">5</option>
                                <option value="6">6</option>
                                <option value="7">7</option>
                                <option value="8">8</option>
                                <option value="9">9</option>
                                <option value="10">10</option>
                                <option value="11">11 - 15</option>
                                <option value="15">15 - 20</option>
                                <option value="20">20 - 25</option>
                                <option value="25">25 - 30</option>
                                <option value="30">30 - 40</option>
                                <option value="40">40 - 50</option>
                                <option value="50">50 - 75</option>
                                <option value="75">75 - 100</option>
                            </select>
                            -->
                            <input type="text" name="NumberOfGuests" maxlength="3" onblur="validateNumberOfGuests();" onchange="showPricing(); " /><br />
                            <span name="ValNumberOfGuests" style="display: none; color: Red">Invalid number</span>
                            <span name="ValNumberOfGuestsRequired" style="display: none; color: Red">Number of guests required</span>
                        </td>
                        <td style="color: Red">*</td>
                    </tr>
                    <tr>
                        <td>Average Age Of Guests:</td>
                        <td>
                            <input type="text" name="AgeOfGuests" maxlength="100" />
                        </td>
                        <td/>
                    </tr>
                    <tr>
                        <td>Name For Your Party:</td>
                        <td>
                            <input type="text" name="PartyName" maxlength="100" />
                        </td>
                        <td />
                    </tr>
                    <tr>
                        <td>Your Email Address:</td>
                        <td>
                            <input type="text" name="Email" maxlength="128" /><br />
                            <span name="ValEmail" style="display: none; color: Red">Invalid email address</span>
                            <span name="ValEmailRequired" style="display: none; color: Red">Email address required</span>
                        </td>
                        <td style="color: Red">*</td>
                    </tr>
                    <tr>
                        <td>Please Confirm Email:</td>
                        <td>
                            <input type="text" name="ConfirmEmail" maxlength="128" /><br />
                            <span name="ValEmailMatch" style="display: none; color: Red">Email address does not match</span>
                        </td>
                        <td/>
                    </tr>
                    <tr>
                        <td>Contact Phone Number:</td>
                        <td>
                            <input type="text" name="Phone" maxlength="20" /><br />
                            <span name="ValPhoneRequired" style="display: none; color: Red">Phone number required</span>
                            <span name="ValPhoneFormat" style="display: none; color: Red">Invalid format</span>
                        </td>
                        <td style="color: Red">*</td>
                    </tr>
           
                    <tr>
                        <td>Event Type:</td>
                        <td>
                            <select name="EventTypeId">
                                <option value=""></option>
                                <% BuildEventTypeOptions %>
                            </select>
                        </td>
                        <td/>
                    </tr>
                    <tr>
                        <td>Comments:</td>
                        <td colspan="2">
                            <span name="charcounter">(1000 characters available)</span>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <textarea name="UserComments" style="width: 100%" maxlength="1000" onkeyup="handleMaxLength()"></textarea>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="2" style="text-align: center">
                            <input id="btnSave" type="button" value="Register" onclick="saveRegistration()" ondblclick="return false;" />
                            &nbsp;
                            <input id="btnCancel" type="button" value="Cancel" onclick="window.location = 'calendar.asp';" />
                        </td>
                    </tr>
                </table>
            </div>

            <!-- footer -->
            <div data-role="footer" class="center" style="height: 55px;">
				<h4 style="position: relative; bottom: 8px;">
                    <span class="social_up">Like us </span>
                    <a href="http://m.facebook.com/gatsplat"><img src="../content/images/mobile/facebook.png" alt="facebook icon" /></a>
                    <a href="http://mobile.twitter.com/gatsplat"><img src="../content/images/mobile/twitter.png" alt="twitter icon"/></a>
                    <span class="social_up"> Follow us </span>
                </h4>
			</div>
        </div>
    </body>
</html>
<%

    sub BuildEventTypeOptions()
        dim types
        set types = new EventTypeCollection

        types.LoadAllActive

        for i = 0 to types.Count - 1
            dim myType
            set myType = types(i)

            response.Write "<option value='" & myType.EventTypeId & "'>" & myType.Description & "</option>"
        next
    end sub

    sub BuildTimeOptions(OpenTime, CloseTime)
        dim currentTime, endTime
        currentTime = OpenTime
        endTime = FormatDateTime(DateAdd("h", -2, GetTimeString(CloseTime)), vbShortTime)

        while currentTime <= endTime
            dim myEvents
            set myEvents = new ScheduledEventCollection

            myEvents.LoadByDateAndTime Request.Form("selectedDate"), currentTime

            if myEvents.TotalPatrons <= APP_SETTING_MAXREGISTERPERTIME then
                response.Write "<option value='" & currentTime & "'>" & GetTimeString(currentTime) & "</option>"
            end if

            currentTime = FormatDateTime(DateAdd("n", 30, GetTimeString(currentTime)), vbShortTime)
        wend
    end sub
%>