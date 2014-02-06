<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="common/EventIncludeList2.asp" -->
<%
    dim myDate
    set myDate = new AvailableDate
    myDate.Load(CDate(Request.QueryString("dt")))

%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title></title>

        <script type="text/javascript" src="content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="content/js/common.js"></script>
        <script type="text/javascript">
            String.prototype.trim = function () {
                return this.replace(/^\s*/, "").replace(/\s*$/, "");
            }

            function closeWindow() {
                parent.$.fancybox.close();
            }

            function saveRegistration() {
                $("#btnSave").attr("disabled", "disabled");
                if (validateForm()) {
                    var url = 'common/eventajax2.asp?act=0&dt=' + escape($("span[name=Date]").html()) +
                                                    '&st=' + escape($("select[name=Time]").val()) +
                                                    '&np=' + escape($("select[name=NumberOfGuests]").val()) +
                                                    '&ap=' + escape($("input[name=AgeOfGuests]").val().trim()) +
                                                    '&em=' + escape($("input[name=Email]").val().trim()) +
                                                    '&ph=' + escape($("input[name=Phone]").val().trim()) +
                                                    '&nm=' + escape($("input[name=Name]").val().trim()) +
                                                    '&ev=' + escape($("select[name=EventTypeId]").val()) +
                                                    '&pn=' + escape($("input[name=PartyName]").val().trim()) +
                                                    '&uc=' + escape($("textarea[name=UserComments]").val().trim()) +
                                                    '&gs=' + escape($("select[name=GunSize]").val());

                    $.ajax({
                        url: url,
                        success: function (data) {
                            if ($("input[name=Invite]").checked) {
                                parent.showSuccess();
                            } else {
                                parent.showGuestList(data);
                            }
                        },
                        error: function () { alert("problems encountered"); }
                    });
                } else {
                    $("#btnSave").removeAttr("disabled");                    
                }
            }

            function showPricing() {
                var numUsers = $("select[name=NumberOfGuests]").val()
                
                if (numUsers.length == 0) {
                    $("#normalRateRow").css("display", "none");
                    $("#reducedRateRow").css("display", "none");
                } else {
                    var dt = new Date($("span[name=Date]").html());

                    if (dt.getDay() == 5 || numUsers >= 10) {
                        $("#normalRateRow").css("display", "none");
                        $("#reducedRateRow").css("display", "table-row");
                    } else {
                        $("#normalRateRow").css("display", "table-row");
                        $("#reducedRateRow").css("display", "none");
                    }
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
                        valid = false;
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

                return valid;
            }
        </script>
    </head>
    <body>
        <p style="text-align: center">Please Completely Fill Out The Registration Form</p>
    
        <table cellpadding="2" cellspacing="3" border="0" style="margin: 0 auto;">
            <tr>
                <td>Date:</td>
                <td>
                    <span name="Date"><%= Request.QueryString("dt") %></span>
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
                <td> </td>
            </tr>
            <tr><td colspan="3" align="center"><span style="color: #990000">Note: If a time is not available on the drop down, it means that <br />
            time is booked.  Simply choose a time before or after that. <br /> Any shown times are available to start.</span></td>
            </tr>
            <tr>
                <td>Approximate Number Of Guests:</td>
                <td>
                    <select name="NumberOfGuests" onchange="showPricing()">
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
                </td>
                <td/>
            </tr>
            <tr id="normalRateRow" style="display: none">
                <td colspan="3" align="center"><span style="color: #990000">Based on your reservation the pricing will be $34.99 per person.</span></td>
            </tr>
            <tr id="reducedRateRow" style="display: none">
                <td colspan="3" align="center"><span style="color: #990000">Based on your reservation the pricing will be $29.99 per person.</span></td>
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
                    <span name="ValPhoneFormat" style="display: none; color: Red">Invalid phone number</span>
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
                <td>Gun Size:</td>
                <td>
                    <select name="GunSize">
                        <option value=""></option>
                        <option value="Not Sure">Not Sure</option>
                        <option value="Smaller 50 Caliber">Smaller 50 Caliber</option>
                        <option value="Standard 68 Caliber">Standard 68 Caliber</option>
                    </select>
                </td>
            </tr>
            <tr>
                <td colspan="3" align="center">
                    <span style="color: #990000">Note: 50 caliber guns are normally for age 10 and under.</span>
                </td>
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
                <td>Create Invites For Guests:</td>
                <td colspan="2">
                    <input type="checkbox" name="Invite" />
                </td>
            </tr>
            <tr>
                <td colspan="2" style="text-align: center">
                    <input id="btnSave" type="button" value="Register" onclick="saveRegistration()" ondblclick="return false;" />
                    &nbsp;
                    <input id="btnCancel" type="button" value="Cancel" onclick="closeWindow();" />
                </td>
            </tr>
        </table>
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

            myEvents.LoadByDateAndTime Request.QueryString("dt"), currentTime

            if myEvents.TotalPatrons <= APP_SETTING_MAXREGISTERPERTIME then
                response.Write "<option value='" & currentTime & "'>" & GetTimeString(currentTime) & "</option>"
            end if

            currentTime = FormatDateTime(DateAdd("n", 30, GetTimeString(currentTime)), vbShortTime)
        wend
    end sub
%>