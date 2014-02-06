<!DOCTYPE html>
<!--#include file="../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForAdminOnly

    dim myEvent, myDate, previewDataTemplate
    set myEvent = new ScheduledEvent
    set myDate = new AvailableDate
    previewDataTemplate = False

    if Request.QueryString(QUERYSTRING_VAR_ID) = 0 then
        myEvent.EventId = 0
        myDate.Load(FormatDateTime(Now(), vbShortDate))
    else
        myEvent.Load(Request.QueryString(QUERYSTRING_VAR_ID))
        myDate.Load(myEvent.SelectedDate)
        if cint(myEvent.EventId) = cint(Settings(SETTING_DEFAULT_EVENTID)) then
            previewDataTemplate = True
        end if
    end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title></title>
        <link type="text/css" rel="stylesheet" href="../../content/css/jquery.ui.all.css" />

        <script type="text/javascript" src="../../content/js/common.js"></script>
        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.ui.core.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.ui.widget.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.ui.datepicker.min.js"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                setNumberOfGuests();
                setStartTime();

                $("input[name=Date]").datepicker();

                window.parent.resizeFancyBox();
            });

            function cancelRegistration() {
                if (confirm("Are you sure?")) {
                    var url = '../../ajax/event.asp?act=4&id=' + escape($("input[name=EventId]").val()) + 
                                                  '&ss=<%=PAYMENTSTATUS_CANCELLED_ADMIN%>&rs=' + escape($("input[name=Resend]").is(":checked"));

                    $.ajax({
                        url: url,
                        success: function () {
                            reloadParent();
                            closeWindow();
                        },
                        error: function () { alert('problems encountered'); }
                    });
                }
            }

            function closeWindow() {
                parent.$.fancybox.close();
            }

            function handleCommentsMaxLength(controlFieldName, counterFieldName) {
                var commentField = $("textarea[name=" + controlFieldName + "]");

                if (commentField.val().length <= commentField.attr("maxlength")) {
                    var remains = commentField.attr("maxlength") - commentField.val().length;
                    $("span[name=" + counterFieldName + "]").html("(" + remains + " characters available)");
                    return true;
                } else {
                    return false;
                }
            }

            function reloadParent() {
                <% if previewDataTemplate = False then %>
                if ('<%= Request.QueryString(QUERYSTRING_VAR_VIEWTYPE)%>' == 'day') {
                    parent.selectDay($("input[name=Date]").val());
                } else {
                    parent.showAll();
                } 
                <% end if %>
            }

            function saveRegistration() {
                if (validateForm()) {
                    var url = '../../ajax/event.asp?act=3&id=' + $("input[name=EventId]").val() +
                                                    '&dt=' + escape($("input[name=Date]").val()) +
                                                    '&st=' + escape($("select[name=Time]").val()) +
                                                    '&np=' + escape(parseInt($("input[name=NumberOfGuests]").val())) +
                                                    '&ap=' + escape($("input[name=AgeOfGuests]").val()) +
                                                    '&em=' + escape($("input[name=Email]").val()) +
                                                    '&ph=' + escape($("input[name=Phone]").val()) +
                                                    '&nm=' + escape($("input[name=Name]").val()) +
                                                    '&ev=' + escape($("select[name=EventTypeId]").val()) +
                                                    '&pn=' + escape($("input[name=PartyName]").val()) +
                                                    '&uc=' + escape($("textarea[name=UserComments]").val()) +
                                                    '&ac=' + escape($("textarea[name=AdminComments]").val()) +
                                                    '&rs=' + escape($("input[name=Resend]").is(":checked"));

                    $.ajax({
                        url: url,
                        success: function () {
                            reloadParent();
                            closeWindow();
                        },
                        error: function () { alert("problems encountered"); }
                    });
                }
            }

            function setNumberOfGuests() {
                $("select[name=NumberOfGuests]").val("<%= myEvent.NumberOfPatrons %>");
            }

            function setStartTime() {
                $("select[name=Time]").val("<%= GetMilitaryTime(myEvent.StartTime) %>");
            }

            //#region validation
            function validateContactName() {
                var valid = true;
                var name = $("input[name=Name]").val();

                if (name.length == 0) {
                    $("span[name=ValName]").css("display", "inline");
                    valid = false;
                } else {
                    $("span[name=ValName]").css("display", "none");
                }

                return valid;
            }

            function validateEmail() {
                var valid = true;
                var email = $("input[name=Email]").val().toLowerCase();
                var confirm = $("input[name=ConfirmEmail]").val().toLowerCase();
                var emailRegExp = /^([a-zA-Z0-9_\.\-\+])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;

                if ($("select[name=Time]").val().length == 0) {
                    $("span[name=ValTime]").css("display", "inline");
                    valid = false;
                } else {
                    $("span[name=ValTime]").css("display", "none");
                }

                if (email.length > 0) {
                    if (emailRegExp.test(email)) {
                        $("span[name=ValEmail]").css("display", "none");
                        $("span[name=ValEmailRequired]").css("display", "none");

                        //if (email != confirm) {
                        //    $("span[name=ValEmailMatch]").css("display", "inline");
                        //    valid = false;
                        //} else {
                            $("span[name=ValEmailMatch]").css("display", "none");
                        //} // confirm email matching
                    } else {
                        $("span[name=ValEmail]").css("display", "inline");
                        valid = false;
                    } // regExp test
                } else {
                    $("span[name=ValEmailRequired]").css("display", "inline");
                    valid = false;
                } // length test

                return valid;
            }

            function validateForm() {
                //return (validateEmail() & validatePhoneNumber() & validateContactName() & validateNumberOfGuests());

                // we are having to play games here
                var valid = !!(validateEmail() & validatePhoneNumber() & validateContactName() & validateNumberOfGuests() & 
                               validateRequiredCustomFields());

                if (valid) {
                    // if the user is using chrome, when the button is disabled, it will not submit
                    // however, submitting the form here in IE and FF will cause the form to also submit
                    // when execution of this function has ceased because we are returning true
                    $("#btnSubmit").attr("disabled", "disabled");
                    $('#myForm').submit();
                } else {
                    $("#btnSubmit").removeAttr("disabled");
                }

                // always return false to keep from double posting
                //return valid;
                return false;
            }

            function validateNumberOfGuests() {
                var valid = true;
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

            function validatePhoneNumber() {
                var valid = true;
                var phoneRegExp = /^\(?(\d{3})\)?[- ]?(\d{3})[- ]?(\d{4})$/;
                var phone = $("input[name=Phone]").val();

                if (phone.length == 0) {
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

                return valid;
            }
     
            function validateRequiredCustomFields() {
                var valid = true;
                var fields = $(".requiredCustom");

                for (i = 0; i < fields.length; i++) {
                    var fieldValidator = $("#" + fields.eq(i).data("reqvalidatorid"));
                    fieldValidator.css("display", "none");

                    if (fields.eq(i).is(":checkbox") && fields.eq(i).is(":checked") == false) {
                        valid = false;
                        fieldValidator.css("display", "block");
                    } else if (fields.eq(i).val().length == 0) {
                        valid = false;
                        fieldValidator.css("display", "block");
                    }
                }

                return valid;
            }
            //#endregion
        </script>
    </head>
    <body>
        <p style="text-align: center">Please Completely Fill Out The Registration Form</p>
    
        <form id="myForm" action="SaveRegistration.asp" method="post" data-ajax="false">
            <table cellpadding="2" cellspacing="3" border="0" style="margin: 0 auto;">
            <tr>
                <td>Date:</td>
                <td>
                    <input type="hidden" name="EventId" value="<%= myEvent.EventId %>" />
                    <input type="text" name="Date" value="<%= FormatDateTime(myDate.SelectedDate, vbShortDate) %>" />
                </td>
                <td colspan="2"/>
            </tr>
             <tr>
                <td>Your Name:</td>
                <td>
                    <input type="text" name="Name" maxlength="100" onblur="validateContactName()" value="<%= myEvent.ContactName %>" /><br />
                    <span name="ValName" style="display: none; color: Red">Name is required</span>
                </td>
                <td style="color: Red">*</td>
                <td />
            </tr>
            <tr>
                <td>Your Email Address:</td>
                <td>
                    <input type="text" name="Email" maxlength="128" onblur="validateEmail()" value="<%= myEvent.ContactEmailAddress %>" /><br />
                    <span name="ValEmail" style="display: none; color: Red">Invalid email address</span>
                    <span name="ValEmailRequired" style="display: none; color: Red">Email address required</span>
                </td>
                <td style="color: Red">*</td>
                <td />
            </tr>
            <tr style="display: none">
                <td>Please Confirm Email:</td>
                <td>
                    <input type="text" name="ConfirmEmail" maxlength="128" onblur="validateEmail()" value="<%= myEvent.ContactEmailAddress %>" /><br />
                    <span name="ValEmailMatch" style="display: none; color: Red">Email address does not match</span>
                </td>
                <td colspan="2"/>
            </tr>
            <tr>
                <td>What Time Do you Want to Start?:</td>
                <td>
                    <select name="Time">
                        <% BuildTimeOptions %>
                    </select><br />
                    <span name="ValTime" style="display: none; color: Red">Start time is required</span>
                </td>
                <td style="color: Red">*</td>
                <td />
            </tr>
            <tr>
                <td>Approximate Number Of Guests:</td>
                <td>
                    <input type="text" name="NumberOfGuests" maxlength="3" onblur="validateNumberOfGuests()" value="<%= myEvent.NumberOfPatrons %>" /><br />
                    <span name="ValNumberOfGuests" style="display: none; color: Red">Invalid number</span>
                    <span name="ValNumberOfGuestsRequired" style="display: none; color: Red">Number of guests required</span>
                </td>
                <td style="color: Red">*</td>
                <td />
            </tr>
            <tr>
                <td>Average Age Of Guests:</td>
                <td>
                    <input type="text" name="AgeOfGuests" maxlength="100" value="<%= myEvent.AgeOfPatrons %>" />
                </td>
                <td colspan="2"/>
            </tr>
            <tr>
                <td>Name For Your Party:</td>
                <td>
                    <input type="text" name="PartyName" maxlength="100" value="<%= myEvent.PartyName %>" />
                </td>
                <td colspan="2" />
            </tr>
            <tr>
                <td>Contact Phone Number:</td>
                <td>
                    <input type="text" name="Phone" maxlength="20" onblur="validatePhoneNumber()" value="<%= myEvent.ContactPhone %>" /><br />
                    <span name="ValPhoneRequired" style="display: none; color: Red">Phone number required</span>
                    <span name="ValPhoneFormat" style="display: none; color: Red">Invalid phone number</span>
                </td>
                <td style="color: Red">*</td>
                <td />
            </tr>
           
            <% BuildCustomFields myEvent.EventId %>

            <tr>
                <td>User Comments:</td>
                <td colspan="3"  style="text-align: right">
                    <span name="userCharCounter">(1000 characters available)</span>
                </td>
                <td />
            </tr>
            <tr>
                <td colspan="4">
                    <textarea name="UserComments" style="width: 100%" maxlength="1000" onkeyup="handleCommentsMaxLength('UserComments', 'userCharCounter')"><%= myEvent.UserComments %></textarea>
                </td>
            </tr>
            <% if previewDataTemplate = False then %>
            <tr>
                <td>Admin Comments:</td>
                <td colspan="3" style="text-align: right">
                    <span name="adminCharCounter">(1000 characters available)</span>
                </td>
            </tr>
            <tr>
                <td colspan="4">
                    <textarea name="AdminComments" style="width: 100%" maxlength="1000" onkeyup="handleCommentsMaxLength('AdminComments', 'adminCharCounter')"><%= myEvent.AdminComments %></textarea>
                </td>
            </tr>
            <tr>
                <td colspan="4">
                    Resend Email&nbsp;
                    <input type="checkbox" name="Resend" />
                </td>
            </tr>
            <% end if %>
            <tr>
                <td colspan="4" style="text-align: center">
                    <button id="btnSubmit" type="submit" onclick="return validateForm();" ondblclick="return false;">Save</button>
                    &nbsp;
                    <button type="button" onclick="closeWindow();">Cancel</button>
                </td>
            </tr>
            <tr>
                <td colspan="4" style="text-align: center">
                    <% if previewDataTemplate = False then %>
                        <button type="button" onclick="cancelRegistration()">Delete This Record</button>
                    <% end if %>
                </td>
            </tr>
        </table>
        </form>
    </body>
</html>
<%

   sub BuildCustomFields(EventId)
        dim fields
        set fields = new RegistrationCustomFieldCollection

        fields.LoadAll
    
        for i = 0 to fields.Count - 1
            response.Write "<tr>" & vbCrLf
    
            select case fields(i).CustomFieldDataType.CustomFieldDataTypeId
                case FIELDTYPE_SHORTTEXT
                    BuildCustomControlTextField fields(i), EventId
                case FIELDTYPE_LONGTEXT
                    BuildCustomControlTextArea  fields(i), EventId
                case FIELDTYPE_YESNO
                    BuildCustomControlCheckBox fields(i), EventId
                case FIELDTYPE_OPTIONS
                    BuildCustomControlDropDownList fields(i), EventId
            end select
                
            response.Write "</tr>" & vbCrLf
        next
    end sub

    sub BuildCustomControlCheckBox(CurrentField, EventId)
        dim currentValue : currentValue = CurrentField.GetFormValue(EventId)
        dim fieldName : fieldName = "CustomField" & CurrentField.RegistrationCustomFieldId
        dim hasAddtionalInfo : hasAddtionalInfo = cbool(len(CurrentField.AdditionalInformation) > 0)
        dim hasNotes : hasNotes = cbool(len(CurrentField.Notes) > 0)
        dim isRequired : isRequired = CurrentField.Required
        dim validatorName : validatorName = "val" & fieldName

        with response
            .Write "<td>" & CurrentField.Name & ":</td>" & vbCrLf
            .Write "<td>" & vbCrLf
            .Write "<input type=""checkbox"" name=""" & fieldName & """ id=""" & fieldName & """ " & iif(cbool(currentValue), "checked=""checked""", "") & " "

            .Write "class="""
            if isRequired then .Write " requiredCustom"
            .Write """ "

            if isRequired then .Write "data-reqValidatorId=""" & validatorName & """"

            .Write "/><br/>" & vbCrLF

            if isRequired then .Write "<span id=""" & validatorName & """ style=""color: red; display: none"">Required</span>" & vbCrLf

            .Write "</td>" & vbCrLf

            if isRequired then .Write "<td><span style=""color: red"">*</span></td>" & vbCrLf else .Write "<td/>" & vbCrLf
            .Write "<td />" & vbCrLf
        end with
    end sub

    sub BuildCustomControlDropDownList(CurrentField, EventId)
        dim currentValue : currentValue = CurrentField.GetFormValue(EventId)
        dim fieldName : fieldName = "CustomField" & CurrentField.RegistrationCustomFieldId
        dim hasAddtionalInfo : hasAddtionalInfo = cbool(len(CurrentField.AdditionalInformation) > 0)
        dim hasNotes : hasNotes = cbool(len(CurrentField.Notes) > 0)
        dim isRequired : isRequired = CurrentField.Required
        dim validatorName : validatorName = "val" & fieldName

        with response
            .Write "<td>" & CurrentField.Name & ":</td>" & vbCrLf
            .Write "<td>" & vbCrLf
            .Write "<select name=""" & fieldName & """ id=""" & fieldName & """  "

            .Write "class="""
            if isRequired then .Write " requiredCustom"
            .Write """ "

            if isRequired then .Write "data-reqValidatorId=""" & validatorName & """"

            .Write ">" & vbCrLf

            .Write "<option></option>" & vbCrLf

            CurrentField.LoadOptions
            dim options
            set options = CurrentField.Options

            for j = 0 to options.Count - 1
                .Write "<option value=""" & options(j).Sequence & """ " & iif(CStr(options(j).Sequence) = currentValue, "selected=""selected""", "") & ">" & options(j).Text & "</option>" & vbCrLf
            next

            .Write "</select><br/>" & vbCrLF

            if isRequired then .Write "<span id=""" & validatorName & """ style=""color: red; display: none"">Required</span>" & vbCrLf

            .Write "</td>" & vbCrLf
            if isRequired then 
                .Write "<td style=""color: Red"">*</td>"  & vbCrLf
            else
                .Write "<td/>" & vbCrLf
            end if
            .Write "<td/>" & vbCrLf
        end with
    end sub 

    sub BuildCustomControlTextArea(CurrentField, EventId)
        dim currentValue : currentValue = CurrentField.GetFormValue(EventId)
        dim fieldName : fieldName = "CustomField" & CurrentField.RegistrationCustomFieldId
        dim hasAddtionalInfo : hasAddtionalInfo = cbool(len(CurrentField.AdditionalInformation) > 0)
        dim hasNotes : hasNotes = cbool(len(CurrentField.Notes) > 0)
        dim isRequired : isRequired = CurrentField.Required
        dim validatorName : validatorName = "val" & fieldName

        with response
            .Write "<td>" & CurrentField.Name & ":</td>" & vbCrLf
            .Write "<td/>" & vbCrLf
            if isRequired then 
                .Write "<td style=""color: Red"">*</td>"  & vbCrLf
                .Write "<td/>" & vbCrLf
            else
                .Write "<td colspan=""3""/>" & vbCrLf
            end if
            .Write "</tr>" & vbCrLf
            .Write "<tr>" & vbCrLf
            .Write "<td colspan=""4"">" & vbCrLf

            .Write "<textarea id=""" & FieldName & """ name=""" & fieldName & """ style=""width: 100%"" "

            .Write "class="""
            if isRequired then .Write " requiredCustom"
            .Write """ "

            if isRequired then .Write "data-reqValidatorId=""" & validatorName & """"

            .Write ">" & currentValue & "</textarea><br/>" & vbCrLf
            
            if isRequired then .Write "<span id=""" & validatorName & """ style=""color: red; display: none"">Required</span>" & vbCrLf

            .Write "</td>" & vbCrLf
        end with
    end sub

    sub BuildCustomControlTextField(CurrentField, EventId)
        dim currentValue : currentValue = CurrentField.GetFormValue(EventId)
        dim fieldName : fieldName = "CustomField" & CurrentField.RegistrationCustomFieldId
        dim hasAddtionalInfo : hasAddtionalInfo = cbool(len(CurrentField.AdditionalInformation) > 0)
        dim hasNotes : hasNotes = cbool(len(CurrentField.Notes) > 0)
        dim isRequired : isRequired = CurrentField.Required
        dim validatorName : validatorName = "val" & fieldName

        with response
            .Write "<td>" & CurrentField.Name & ":</td>" & vbCrLf
            .Write "<td>" & vbCrLf
            .Write "<input type=""text"" name=""" & fieldName & """ id=""" & fieldName & """ "

            .Write "class="""
            if isRequired then .Write " requiredCustom"
            .Write """ "

            if isRequired then .Write "data-reqValidatorId=""" & validatorName & """"

            .Write " value=""" & currentValue & """/><br/>" & vbCrLF

            if isRequired then .Write "<span id=""" & validatorName & """ style=""color: red; display: none"">Required</span>" & vbCrLf

            .Write "</td>" & vbCrLf
            if isRequired then 
                .Write "<td style=""color: Red"">*</td>"  & vbCrLf
            else
                .Write "<td/>" & vbCrLf
            end if
            .Write "<td/>" & vbCrLf
        end with
    end sub

    sub BuildTimeOptions()
        dim currentTime, endTime
        currentTime = FormatDateTime("12:00 am", vbShortTime)
        endTime = FormatDateTime("12:00 am", vbShortTime)

        do 
            response.Write "<option value='" & currentTime & "'>" & GetTimeString(currentTime) & "</option>"
            currentTime = FormatDateTime(DateAdd("n", 15, currentTime), vbShortTime)
        loop until currentTime = endTime 
    end sub
%>