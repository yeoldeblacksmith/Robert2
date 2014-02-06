<!DOCTYPE html>
<!--#include file="../classes/IncludeList.asp" -->
<%
    if cbool(settings(SETTING_MODULE_REGISTRATION)) = false then response.Redirect("disabled.asp")

    dim myDate : set myDate = new AvailableDate
    myDate.Load(CDate(Request.QueryString("dt")))

    dim myEvent : set myEvent = new ScheduledEvent
    myEvent.SelectedDate = myDate.ToDate()
    myEvent.CreateDateTime = Now()
    myEvent.PaymentStatusId = PAYMENTSTATUS_TEMPORARY
    myEvent.Save

%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Event Registration</title>
        <meta name="viewport" content="width=device-width, initial-scale=1"/>
        <meta http-equiv="cache-control" content="max-age=0" /> 
        <meta http-equiv="cache-control" content="no-cache" /> 
        <meta http-equiv="expires" content="0" /> 
        <meta http-equiv="expires" content="Tue, 01 Jan 1980 1:00:00 GMT" /> 
        <meta http-equiv="pragma" content="no-cache" />
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>

        <link type="text/css" rel="stylesheet" href="../content/css/jquery.mobile-1.3.2.css" />
        <link type="text/css" rel="stylesheet" href="../content/css/jquery.fancybox-1.3.4.css" />
        <link type="text/css" rel="stylesheet" href="../content/css/jquery.powertip-light.css" />
        <link type="text/css" rel="Stylesheet" href="../content/css/eventregistration.css?<%=ANTI_CACHE_STRING%>" />
        <link type="text/css" rel="Stylesheet" href="../content/css/waiver.css?<%=ANTI_CACHE_STRING%>" />

        <script type="text/javascript" src="../content/js/common.js?<%=ANTI_CACHE_STRING%>"></script>
        <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.mobile-1.3.2.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.fancybox-1.3.4.pack.js"></script>
        <script type="text/javascript" src="../content/js/jquery.powertip.min.js"></script>
        <script type="text/javascript">
            var bCheckedForDupes = false;
            var paymentCompleted = false;

            String.prototype.trim = function () {
                return this.replace(/^\s*/, "").replace(/\s*$/, "");
            }

            $(document).ready(function () {
                $(".tooltip").powerTip({ mouseOnToPopup: true, placement: 'n', smartPlacement: true });
                //reloadAvailableTimes($("#NumberOfGuests").val());

                calculateTotals();
                $("input:not([type='hidden']):not([type='checkbox'])").textinput({ preventFocusZoom: true });
                $("select").selectmenu({ preventFocusZoom: true });
            });

            function addDetailRowToCart(desc, qty, price, extPrice) {
                var row = $("<tr>").appendTo($("#tblPaymentDetails"));
                $("<td>").text(desc).appendTo(row);
                $("<td>").text(qty).appendTo(row);
                $("<td>").text(price.toFixed(2)).appendTo(row);
                $("<td>").text(extPrice.toFixed(2)).appendTo(row);
            }

            function addHeaderRowToCart(){
                var row = $("<tr>").appendTo($("#tblPaymentDetails"));
                $("<th>").text("Description").appendTo(row);
                $("<th>").text("Qty").appendTo(row);
                $("<th>").text("Price").appendTo(row);
                $("<th>").text("Ext. Price").appendTo(row);
            }

            function addTotalRowsToCart(total) {
                $("<tr><td colspan='4'><hr/></td></tr>").appendTo($("#tblPaymentDetails"));
                var row = $("<tr>").appendTo($("#tblPaymentDetails"));
                $("<td>").attr("colspan", "3").text("Total").appendTo(row);
                $("<td>").text(total.toFixed(2)).appendTo(row);
            }

            function calculateTotals() {
                var total = 0;
                var players = 0;

                if($("#NumberOfGuests").val().length > 0) { players = players = parseInt($("#NumberOfGuests").val()); }

                var detailTable = $("#tblPaymentDetails");
                detailTable.empty();

                addHeaderRowToCart();

                if ($("#txtBaseFee").length > 0) {
                    var baseFee = parseFloat($("#txtBaseFee").val());

                    switch (parseInt($("#txtBaseFeeType").val())) {
                        case <%=PAYMENTTYPE_BYEVENT%>:
                            addDetailRowToCart('Registration Deposit', 1, baseFee, baseFee);
                            break;
                        case <%=PAYMENTTYPE_BYPLAYER%>:
                            if (players == 0) {
                                baseFee = 0;
                            } else {                               
                                var extPrice = players * baseFee;
                                addDetailRowToCart('Registration Deposit', players, baseFee, extPrice);

                                baseFee = extPrice;
                            } 

                            break;
                        default:
                            baseFee = 0;
                            break;
                    }

                    total += baseFee;
                }

                $("select[id^=CustomField]").each(function() {
                    if(typeof $(this).data("feefieldid") != 'undefined') {
                        var selected = $(this).find(":selected");

                        if (selected.length > 0 ){
                            var value = selected.data("value");
                            var text = selected.data("text");

                            if(typeof value != 'undefined') {
                                switch($(this).data("feetype")){
                                    case <%=PAYMENTTYPE_BYEVENT%>:
                                        addDetailRowToCart(text, 1, value, value);
                                        total += value;
                                        break;
                                    case <%=PAYMENTTYPE_BYPLAYER%>:
                                        var extPrice = players * value;
                                        addDetailRowToCart(text, players, value, extPrice);
                                        total += extPrice;
                                        break;
                                }
                            }
                        }
                    }
                });

                $("#txtCartTotal").val(total);
                if (detailTable.html().length > 0) {addTotalRowsToCart(total);}
                if (total == 0) {
                    detailTable.css("display", "none");
                    $("#btnSubmit").prev("span").find(".ui-btn-text").html("Register");

                    $("#myForm").attr("action", "save.asp");
                    $("#myForm").attr("target", "_self");
                } else {
                    detailTable.css("display", "table");
                    $("#btnSubmit").prev("span").find(".ui-btn-text").html("Check Out");

                    $("#myForm").attr("action", "cartcheckout.asp");
                    $("#myForm").attr("target", "_blank");
                }

                toggleSubmitButton();
            }

            function closeWindow() {
                window.location.href = '<%= SiteInfo.HomeUrl %>';
            }

            function handleUserMaxLength(length) {
                var commentField = $("#UserComments");

                if (commentField.val().length <= commentField.attr("maxlength")) {
                    var remains = commentField.attr("maxlength") - commentField.val().length;
                    $("#charcounter").html("(" + remains + " characters available)");
                    return true;
                } else {
                    return false;
                }
            }

            function hideAvailableTimeWindow() {
                //$("#popupTimes").popup("close");
                $.fancybox.close();
            }

            function NumberOfGuests_onblur() {
                calculateTotals();

                if (!!validateNumberOfGuests() == false) {
                    $("#btnSubmit").button("disable");
                }
            }

            function pollPayment() {
                $.ajax({
                    url: '../ajax/payment.asp?act=1&id=' + $("#EventId").val(),
                    error: function () { alert("problems encountered polling for payments"); },
                    success: function (results) {
                        if (results == "completed") {
                            paymentCompleted = true;
                        }
                    }
                });
            }

            function redirectHome() {
                window.location.href = '<%= siteInfo.HomeUrl %>';
            }

            function resendConfirmation() {
                var eventId;

                if ($("#tblEvents").css("display") === "table") {
                    eventId = $("input[name=radEventId]:checked").val();
                } else {
                    eventId = $("#txtDuplicateEventId").val();
                }

                $.ajax({
                    url: '../ajax/confirmemail.asp?act=0&ev=' + eventId,
                    error: function () { alert("error encountered sending email"); },
                    async: false
                });
            }

            function resizeFancyBox(){
                $('#fancybox-content').height($('#fancybox-frame').contents().find('div.waiverSection').height() + 5);
                $.fancybox.center();
            }

            function runCounter() {
                if (paymentCompleted == false) {
                    pollPayment(); 
                    setTimeout("runCounter()", 5000);
                } else {
                    var myForm = $("#myForm");
                    myForm.attr("action", "PaymentReceived.asp");
                    myForm.attr("target", "_self");
                    myForm.submit();
                }
            }

            function setStartTime(Time, Display) {
                $("#lblStartTime").html(Display);
                $("#Time").val(Time);

                if ($("#chkTerms").length > 0) {
                    toggleSubmitButton();
                } else {
                    $("#btnSubmit").button("enable");
                }
            }

            function showActionWindow() {
                $.ajax({
                    url: '../ajax/event.asp?act=10&dt=' + escape($("#Date").html()) +
                                                 '&em=' + escape($("#Email").val().trim()),
                    cache: false,
                    dataType: "json",
                    error: function () { alert("error encountered while trying to validate duplicates"); },
                    success: function (result) {
                        if (result.length == 1) {
                            $("#txtDuplicateEventId").val(result[0].EventId);
                        } else {
                            var tbl = $("#tblEvents");

                            for (i = 0; i < result.length; i++) {
                                var row = $("<tr>");
                                var cellSelection = $("<td>");
                                var cellDate = $("<td>").html(result[i].SelectedDate);
                                var cellTime = $("<td>").html(result[i].StartTime);
                                var cellName = $("<td>").html(result[i].PartyName);

                                var selector = $("<input>").attr("type", "radio").attr("name", "radEventId").val(result[i].EventId)
                                if (i == 0) { selector.attr("checked", "checked"); }

                                cellSelection.append(selector);

                                row.append(cellSelection);
                                row.append(cellDate);
                                row.append(cellTime);
                                row.append(cellName);

                                tbl.append(row);
                            }

                            tbl.css("display", "table");
                        }

                        var alertHeight = 295 + ((result.length - 1) * 50);

                        $.fancybox({
                            'content': $("#divDuplicateAction").html(),
                            'showCloseButton': false,
                            'modal': true,
                            'centerOnScroll': true,
                            'width': 450,
                            'height': alertHeight,
                            'autoDimensions': false
                        });
                    },
                    async: false
                });
            }

            function showAdditionalInfo(customField) {
                $.fancybox({
                    'content': $(customField).data("info"),
                    'showCloseButton': true
                });

                //var infoWindow = $("#popupAdditionalInfo");
                //infoWindow.html($(customField).data("info"));

                //infoWindow.popup({
                //    dismissible: true,
                //    history: false,
                //    overlayTheme: "a",
                //    positionTo: "window",
                //    shadow: false,
                //    transition: "flip"
                //});

                //infoWindow.popup("open");
            }

            function showAvailableTimes() {
                if (validateNumberOfGuests()) {
                    var selDate = $("#SelectedDate").val();
                    var playerCnt = $("#NumberOfGuests").val();
                    var timeUrl = "AvailableTimeList.asp?dt=" + selDate + '&np=' + playerCnt;

                    $.fancybox({
                        'type': 'iframe',
                        'href': timeUrl,
                        'showCloseButton': false,
                        'modal': true,
                        'centerOnScroll': true,
                        'width': 300,
                        'padding': 2,
                        'margin': 0,
                        'autoDimensions': false,
                        'autoScale': false,
                        'scrolling': 'no'
                    });

                }

                return false;
            }

            function showConfirmationWindow() {
                $.fancybox({
                    'content': $("#divDuplicateConfirmation").html(),
                    'showCloseButton': false,
                    'modal': true,
                    'centerOnScroll': true,
                    'width': 350,
                    'height': 120,
                    'autoDimensions': false
                });
            }

            function showTerms() {
                var dlg = $("<div />")
                    .attr("data-role", "dialog")
                    .attr("id", "dialog")
                    .attr("data-transition", "flip");
                
                var hdr = $("<div>")
                    .attr("data-role", "header")
                    .attr("data-theme", "b");

                var heading = $("<h2>").html("Terms and Conditions");
                hdr.append(heading);

                dlg.append(hdr);

                var content = $("<div />")
                    .attr("data-role", "content")
                    .append($("#dialogTerms").html());

                dlg.append(content);

                dlg.appendTo($.mobile.pageContainer);

                // show the dialog programmatically
                $.mobile.changePage(dlg, { role: "dialog" });
            }

            function toggleSubmitButton() {
                //if ($("#chkAgreement").is(":checked")) {
                //if ($("#chkTerms").is(":checked") && $("#Time").val().length > 0) {
                if ($("#chkTerms").is(":checked")) {
                    $("#btnSubmit").button("enable");
                } else {
                    $("#btnSubmit").button("disable");
                }
            }

            //#region Validation
            function validateContactName() {
                var valid = true;

                var name = $("#Name").val();

                if (name.trim().length == 0) {
                    $("#ValName").css("display", "inline");
                    valid = false;
                } else {
                    $("#ValName").css("display", "none");
                }

                return valid;
            }

            function validateDuplicate() {
                var valid = true;

                if (bCheckedForDupes == false) {
                    $.ajax({
                        url: '../ajax/event.asp?act=10&dt=' + escape($("#SelectedDate").val()) +
                                                     '&em=' + escape($("#Email").val().trim()),
                        cache: false,
                        dataType: "json",
                        error: function () { alert("error encountered while trying to validate duplicates"); },
                        success: function (result) {
                            if (result.length > 0) {
                                $.fancybox({
                                    'content': $("#divDuplicateWarning").html(),
                                    'showCloseButton': false,
                                    'modal': true,
                                    'centerOnScroll': true,
                                    'width': 350,
                                    'height': 275,
                                    'autoDimensions': false
                                });

                                valid = false;
                            }
                        },
                        async: false
                    });

                }

                return valid;
            }

            function validateEmail() {
                var valid = true;
                //var email = $("input[name=Email]").val().toLowerCase();
                var email = $("#Email").val().toLowerCase();
                var confirm = $("#ConfirmEmail").val().toLowerCase();
                var emailRegExp = /^([a-zA-Z0-9_\.\-\+])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;

                if (email.length > 0) { //if (email.trim().length > 0) {
                    if (emailRegExp.test(email)) {
                        $("#ValEmail").css("display", "none");
                        $("#ValEmailRequired").css("display", "none");

                        //if (email != confirm) {
                        //    $("#ValEmailMatch").css("display", "inline");
                        //    valid = false;
                        //} else {
                        $("#ValEmailMatch").css("display", "none");
                        //} // confirm email matching
                    } else {
                        $("#ValEmail").css("display", "inline");
                        valid = false;
                    } // regExp test
                } else {
                    $("#ValEmailRequired").css("display", "inline");
                    valid = false;
                } // length test

                return valid;
            }

            function validateForm() {
                // we are having to play games here
                var valid = !!(validateDuplicate() & validateEmail() & validatePhoneNumber() &
                               validateContactName() & validateNumberOfGuests() & validatePartyName() &
                               validateTime() & validateRequiredCustomFields());

                if (valid) {
                    // if the user is using chrome, when the button is disabled, it will not submit
                    // however, submitting the form here in IE and FF will cause the form to also submit
                    // when execution of this function has ceased because we are returning true
                    //$("#btnSubmit").attr("disabled", "disabled");
                    $('#myForm').submit();

                    if ($("#myForm").attr("action") == "cartcheckout.asp") {runCounter();} // this will only run for check outs 
                } else {
                    //$("#btnSubmit").removeAttr("disabled");
                }

                // always return false to keep from double posting
                //return valid;
                return false;
            }

            function validateNumberOfGuests() {
                var valid = true;
                var guestCnt = $("#NumberOfGuests").val();

                $("#ValNumberOfGuests").css("display", "none");
                $("#ValNumberOfGuestsRequired").css("display", "none");
                $("#ValNumberOfGuests").css("display", "none");
                $("#valNumberOfGuestsMin").css("display", "none");

                if (guestCnt.trim().length == 0) {
                    $("#ValNumberOfGuestsRequired").css("display", "inline");
                    valid = false;
                } else {
                    if (!isInt(guestCnt)) {
                        $("#ValNumberOfGuests").css("display", "inline");
                        valid = false;
                    } else {
                        if (guestCnt <= 0) {
                            $("#ValNumberOfGuests").css("display", "inline");
                            valid = false;
                        } else {

                            var minReq = <%= iif(isnull(Settings(SETTING_REGISTRATION_MINPLAYERS)) or len(Settings(SETTING_REGISTRATION_MINPLAYERS)) = 0, 1, Settings(SETTING_REGISTRATION_MINPLAYERS)) %>;
                            if (guestCnt < minReq) {
                                $("#valNumberOfGuestsMin").css("display", "inline");
                                valid = false;
                            }
                        }
                    }
                }

                return valid;
            }

            function validatePartyName() {
                var valid = true;
                var partyName = $("#PartyName").val();

                if (partyName.trim().length == 0) {
                    $("#ValPartyNameRequired").css("display", "inline");
                    valid = false;
                } else {
                    $("#ValPartyNameRequired").css("display", "none");
                }


                return valid;
            }

            function validatePhoneNumber() {
                var valid = true;

                var phoneRegExp = /^\(?(\d{3})\)?[- ]?(\d{3})[- ]?(\d{4})$/;

                var phone = $("#Phone").val();

                if (phone.trim().length == 0) {
                    $("#ValPhoneRequired").css("display", "inline");
                    valid = false;
                } else {
                    $("#ValPhoneRequired").css("display", "none");

                    if (phoneRegExp.test(phone)) {
                        $("#ValPhoneFormat").css("display", "none");
                    } else {
                        $("#ValPhoneFormat").css("display", "inline");
                        valid = false;
                    }
                }

                return valid;
            }

            function validateRequiredCustomFields() {
                var valid = true;
                var fields = $("input.requiredCustom,select.requiredCustom");

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

            function validateTime() {
                var valid = true;

                if($("#Time").val().length == 0) {
                    valid = false;
                    $("#valTimeReq").css("display", "inline");
                } else {
                    $("#valTimeReq").css("display", "none");
                }

                return valid;
            }
            //#endregion
        </script>
    </head>

    <body>
        <div class="waiverSection blockCenter" style="width: 550px">
            <img src="<%= LOGO_URL %>" style="display: block; margin: auto" />
            <h3 style="text-align: center">Please Completely Fill Out The Registration Form</h3>
    
            <form action="Save.asp" id="myForm" method="post" data-ajax="false">
                <input type="hidden" id="EventId" name="EventId" value="<%= EncodeId(myEvent.EventId) %>" />
                <table cellpadding="2" cellspacing="3" border="0" style="margin: 0 auto;">
                    <tr>
                        <td style="width:145px; text-align: right">Date:</td>
                        <td>
                            <input type="hidden" name="SelectedDate" id="SelectedDate" value="<%= Request.Querystring("dt") %>" />
                            <span id="Date" name="Date"><%= Request.QueryString("dt") %></span>
                        </td>
                        <td colspan="2"/>
                    </tr>
                    <tr>
                        <td style="width:145px; text-align: right">Your Name:</td>
                        <td>
                            <input type="text" id="Name" name="Name" maxlength="100" onBlur="this.value = this.value.trim(); validateContactName();" />
                            <span id="ValName" name="ValName" style="display: none; color: Red">Name is required</span>
                        </td>
                        <td style="color: Red"> &nbsp; *</td>
                        <td />
                    </tr>
                    <tr>
                        <td style="width:145px; text-align: right">Your Email Address:</td>
                        <td>
                            <input type="email" id="Email" name="Email" maxlength="128" onBlur="this.value = this.value.trim(); validateEmail(); validateDuplicate()" onChange="javascript: bCheckedForDupes = false;" />
                            <span id="ValEmail" name="ValEmail" style="display: none; color: Red">Invalid email address</span>
                            <span id="ValEmailRequired" name="ValEmailRequired" style="display: none; color: Red">Email address required</span>
                        </td>
                        <td style="color: Red"> &nbsp; *</td>
                        <td />
                    </tr>
                    <tr style="display: none">
                        <td style="width:145px; text-align: right">Please Confirm Email:</td>
                        <td>
                            <input type="email" id="ConfirmEmail" name="ConfirmEmail" maxlength="128" onBlur="this.value = this.value.trim(); validateEmail()" />
                            <span id="ValEmailMatch" name="ValEmailMatch" style="display: none; color: Red">Email address does not match</span>
                        </td>
                        <td colspan="2"/>
                    </tr>
                    <tr>
                        <td style="width:145px; text-align: right">Number Of Players:</td>
                        <td>
                            <!--<input type="text" id="NumberOfGuests" name="NumberOfGuests" maxlength="3" onBlur="this.value = this.value.trim(); validateNumberOfGuests(); reloadAvailableTimes(this.value)" />-->
                            <input type="text" id="NumberOfGuests" name="NumberOfGuests" maxlength="3" onBlur="this.value = this.value.trim(); NumberOfGuests_onblur()" />
                            <span id="ValNumberOfGuests" name="ValNumberOfGuests" style="display: none; color: Red">Invalid number</span>
                            <span id="ValNumberOfGuestsRequired" name="ValNumberOfGuestsRequired" style="display: none; color: Red">Number of Players required</span>
                            <span id="valNumberOfGuestsMin" style="display: none; color: red">At least <%= Settings(SETTING_REGISTRATION_MINPLAYERS) %> players are required</span>
                        </td>
                        <td style="color: Red"> &nbsp; *</td>
                        <td>
                            <img id="imgLoading" src="../content/images/loading.gif" alt="loading" style="display: none"/>
                        </td>
                    </tr>
                    <tr id="rowTime">
                        <td class="inlineRight" style="width: 145px">Starting Time:</td>
                        <td>
                            <table style="width: 100%">
                                <tr>
                                    <td>
                                        <input type="hidden" name="Time" id="Time" />
                                        <span id="lblStartTime">No time selected</span>
                                    </td>
                                    <td class="right" style="float: right">
                                        <button onclick="return showAvailableTimes()" id="btnTime" data-inline="true" data-mini="true">Choose Time</button>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2">
                                        <span id="valTimeReq" style="display: none; color: red">Start Time Required</span>
                                    </td>
                                </tr>
                            </table>
                        </td>
                        <td colspan="2"/>
                    </tr>
                    <tr>
                        <td style="width:145px; text-align: right">Average Age Of Guests:</td>
                        <td>
                            <input type="text" id="AgeOfGuests" name="AgeOfGuests" maxlength="100" />
                        </td>
                        <td colspan="2"/>
                    </tr>
                    <tr>
                        <td style="width:145px; text-align: right">Name For Your Party:</td>
                        <td>
                            <input type="text" id="PartyName" name="PartyName" maxlength="100" onBlur="this.value = this.value.trim(); validatePartyName()" 
                                class="tooltip" data-powertip="Note: Please use a meaningful name for the party for your guests to find easily while filling out waivers." />
                            <span id="ValPartyNameRequired" style="display: none; color: Red">Party name required</span>
                        </td>
                        <td style="color: Red"> &nbsp; *</td>
                        <td />
                    </tr>
                    <tr>
                        <td style="width:145px; text-align: right">Contact Phone Number:</td>
                        <td>
                            <input type="tel" id="Phone" name="Phone" maxlength="20" onBlur="this.value = this.value.trim(); validatePhoneNumber()" />
                            <span id="ValPhoneRequired" name="ValPhoneRequired" style="display: none; color: Red">Phone number required</span>
                            <span id="ValPhoneFormat" name="ValPhoneFormat" style="display: none; color: Red">Invalid phone number</span>
                        </td>
                        <td style="color: Red"> &nbsp; *</td>
                        <td />
                    </tr>
                    <% BuildCustomFields %>
                    <tr>
                        <td colspan="4" style="max-width: 425px">
                            <%= Settings(SETTING_BLURB_REGISTRATION) %>
                        </td>
                    </tr>
                    <tr>
                        <td style="width:145px; text-align: right">Comments:</td>
                        <td colspan="3" style="text-align: right" >
                            <span id="charcounter" name="charcounter">(1000 characters available)</span>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="4">
                            <textarea id="UserComments" name="UserComments" style="resize: none; width: 98%" maxlength="1000" onKeyUp="handleUserMaxLength()"></textarea>
                        </td>
                    </tr>
                    <%  if true = false then 'if AllowsCancelOrEdit() then %> 
                    <tr>
                        <td colspan="4">
                            <label>
                            <input type="checkbox" id="chkAgreement" onclick="toggleSubmitButton()" data-inline="true" data-theme="e" data-mini="true" />
                            <%= GetCancelEditText() %>
                            </label>
                        </td>
                    </tr>
                    <% end if %>

                    <% if len(Settings(SETTING_BLURB_REGISTRATION_TERMS)) > 0 then %>
                    <tr>
                        <td colspan="4" class="center">
                            <label>
                            <input type="checkbox" id="chkTerms" onclick="toggleSubmitButton()" data-inline="true" data-theme="e" data-mini="true" />
                            By clicking here, you agree to the terms and conditions for this site.
                            </label>

                            <a onclick="javascript: showTerms();" href="#" data-role="none" >Terms and Conditions</a>
                        </td>
                    </tr>
                    <% else %>
                    <tr>
                        <td colspan="4" class="center">
                            <input type="checkbox" id="chkTerms" style="display: none" checked="checked" />
                        </td>
                    </tr>
                    <% end if %>
                    <tr>
                        <td colspan="4">
                            <input type="hidden" id="txtBaseFee" value="<%= FormatNumber(CDbL(Settings(SETTING_REGISTRATION_PAYMENTAMOUNT)),2) %>" />
                            <input type="hidden" id="txtBaseFeeType" value="<%=Settings(SETTING_REGISTRATION_PAYMENTTYPE) %>" />
                            <input type="hidden" id="txtCartTotal" />
                            <table id="tblPaymentDetails" class="blockCenter" style="margin-top: 20px">
                                <tr>
                                    <th>Description</th>
                                    <th>Qty</th>
                                    <th>Price</th>
                                    <th>Ext. Price</th>
                                </tr>
                                <tr>
                                    <td colspan="4">
                                        <hr/>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="3" />
                                    <td class="inlineRight">0.00</td>
                                </tr>
                            </table>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="4" style="text-align: center">
                            <div data-role="controlgroup"  data-type="horizontal">
                                <button type="submit" id="btnSubmit" onclick="return validateForm();" ondblclick="return false;" disabled="disabled" data-theme="b">Register</button>
                                <button type="button" id="btnCancel" onclick="closeWindow();">Cancel</button>
                            </div>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="4" style="text-align: center">
                            <span id="valDuplicate" style="display: none; color: red">A reservation has already been made for this date</span>
                        </td>
                    </tr>
                </table>
            </form>
        </div>
        <br /><br /><br />


        <div id="divDuplicateWarning" style="display: none; width: 350px; height: 350px">
            <h2 style="text-align: center">Another Registration Found</h2>
            <p>We have found another event scheduled with us. Do you wish to modify this event or create an additional event?</p>
            <table style="width: 350px">
                <tr>
                    <td>
                        <button id="btnModify" onclick="javascript: showActionWindow();">Modify Existing</button>
                    </td>
                </tr>
                <tr>
                    <td>
                        <button id="btnCreate" onclick="javascript: $.fancybox.close(); bCheckedForDupes = true;">Create Additional</button>
                    </td>
                </tr>
            </table>
        </div>

        <div id="divDuplicateAction" style="display: none">
            <h2 style="text-align: center">Modify Existing</h2>
            <p>To edit an event, you will need to use the link found in the confirmation email. If you need the email resent, click the resend button below.</p>
            
            <table id="tblEvents" style="display: none; margin: auto">
                <tr>
                    <th/>
                    <th>Date</th>
                    <th>Time</th>
                    <th>Party</th>
                </tr>
            </table>

            <input type="hidden" id="txtDuplicateEventId" />

            <table style="width: 350px; margin: auto">
                <tr>
                    <td>
                        <button id="btnResend" onclick="javascript: resendConfirmation(); showConfirmationWindow(); return false;">Resend Email</button>
                    </td>
                </tr>
                <tr>
                    <td>
                        <button id="btnContinue" onclick="javascript: redirectHome(); return false;">Exit</button>
                    </td>
                </tr>
            </table>
        </div>

        <div id="divDuplicateConfirmation" style="display: none">
            <h2 style="text-align: center">Email Sent</h2>

            <button  onclick="javascript: redirectHome(); return false;" style="width: 330px">Exit</button>
        </div>

        <div id="dialogTerms" style="display: none">
            <%= Settings(SETTING_BLURB_REGISTRATION_TERMS) %>
        </div>

        <div id="popupTimes" style="display: none">
            <iframe style="border: none; padding: 0px"></iframe>
        </div>

        <div id="popupAdditionalInfo" data-role="popup"></div>
    </body>
</html>
<%
    function AllowsCancelOrEdit()
        AllowsCancelOrEdit = (cbool(Settings(SETTING_REGISTRATION_ALLOWCANCELLATION)) and cint(Settings(SETTING_REGISTRATION_MAXDAYS_CANCEL)) > 0) or _
                             (cbool(Settings(SETTING_REGISTRATION_ALLOWEDIT)) and cint(Settings(SETTING_REGISTRATION_MAXDAYS_EDIT)) > 0 )
    end function

    sub BuildCustomFields()
        dim fields
        set fields = new RegistrationCustomFieldCollection

        fields.LoadAll
    
        for i = 0 to fields.Count - 1
            response.Write "<tr>" & vbCrLf
    
            select case fields(i).CustomFieldDataType.CustomFieldDataTypeId
                case FIELDTYPE_SHORTTEXT
                    BuildCustomControlTextField fields(i)
                case FIELDTYPE_LONGTEXT
                    BuildCustomControlTextArea  fields(i)
                case FIELDTYPE_YESNO
                    BuildCustomControlCheckBox fields(i)
                case FIELDTYPE_OPTIONS
                    BuildCustomControlDropDownList fields(i)
            end select
                
            response.Write "</tr>" & vbCrLf
        next
    end sub

    sub BuildCustomControlCheckBox(CurrentField)
        dim fieldName : fieldName = "CustomField" & CurrentField.RegistrationCustomFieldId
        dim hasAddtionalInfo : hasAddtionalInfo = cbool(len(CurrentField.AdditionalInformation) > 0)
        dim hasNotes : hasNotes = cbool(len(CurrentField.Notes) > 0)
        dim isRequired : isRequired = CurrentField.Required
        dim validatorName : validatorName = "val" & fieldName

        with response
            .Write "<td style=""text-align: right"">" & CurrentField.Name & ":</td>" & vbCrLf
            .Write "<td>" & vbCrLf
            .Write "<input type=""checkbox"" name=""" & fieldName & """ id=""" & fieldName & """ "

            .Write "class="""
            if hasNotes then .Write " tooltip"
            if isRequired then .Write " requiredCustom"
            .Write """ "

            if hasNotes then .write "data-powertip=""" & nullablereplace(CurrentField.Notes, """", "&quot;") & """ "
            if isRequired then .Write "data-reqValidatorId=""" & validatorName & """"

            .Write "/><br/>" & vbCrLF

            if isRequired then .Write "<span id=""" & validatorName & """ style=""color: red; display: none"">Required</span>" & vbCrLf

            .Write "</td>" & vbCrLf

            if isRequired then .Write "<td><span style=""color: red"">&nbsp; *</span></td>" & vbCrLf else .Write "<td/>" & vbCrLf
            if hasAddtionalInfo then
                .Write "<td>" & vbCrLf
                .Write "<img src=""../content/images/info.gif"" class=""mouseover"" alt="""" title=""additional information"" border=""0"" data-info=""" & _
                        NullableReplace(currentField.AdditionalInformation, """", "&quot;") & """ onclick=""showAdditionalInfo(this)"" />" & vbCrLf
                .Write "</td>" & vbCrLf
            else
                .Write "<td />" & vbCrLf
            end if
        end with
    end sub

    sub BuildCustomControlDropDownList(CurrentField)
        dim fieldName : fieldName = "CustomField" & CurrentField.RegistrationCustomFieldId
        dim hasAddtionalInfo : hasAddtionalInfo = cbool(len(CurrentField.AdditionalInformation) > 0)
        dim hasNotes : hasNotes = cbool(len(CurrentField.Notes) > 0)
        dim isRequired : isRequired = CurrentField.Required
        dim validatorName : validatorName = "val" & fieldName

        with response
            .Write "<td style=""text-align: right"">" & CurrentField.Name & ":</td>" & vbCrLf
            .Write "<td>" & vbCrLf
            .Write "<select name=""" & fieldName & """ id=""" & fieldName & """  "

            .Write "class="""
            if hasNotes then .Write " tooltip"
            if isRequired then .Write " requiredCustom"
            .Write """ "

            if hasNotes then .write "data-powertip=""" & nullablereplace(CurrentField.Notes, """", "&quot;") & """ "
            if isRequired then .Write "data-reqValidatorId=""" & validatorName & """"
            if CurrentField.HasPayment then .Write "data-feetype=""" & CurrentField.PaymentTypeId & """ data-feefieldid=""" & CurrentField.RegistrationCustomFieldId & """ onchange=""calculateTotals()"""

            .Write ">" & vbCrLf

            .Write "<option></option>" & vbCrLf

            CurrentField.LoadOptions
            dim options
            set options = CurrentField.Options

            for j = 0 to options.Count - 1
                if CurrentField.HasPayment then
                '    dim optionText : optionText = options(j).Text & " (" & GetPaymentText(CurrentField.PaymentTypeId, options(j).Value) & ")"
                '    .Write "<option value=""" & options(j).Sequence & """ data-value=""" & options(j).Value & """ data-text=""" & options(j).Text & """>" & optionText & "</option>" & vbCrLf
                    .Write "<option value=""" & options(j).Sequence & """ data-value=""" & options(j).Value & """ data-text=""" & options(j).Text & """>" & options(j).Text & "</option>" & vbCrLf
                else
                    .Write "<option value=""" & options(j).Sequence & """>" & options(j).Text & "</option>" & vbCrLf
                end if
            next

            .Write "</select>" & vbCrLF

            if isRequired then .Write "<span id=""" & validatorName & """ style=""color: red; display: none"">Required</span>" & vbCrLf

            .Write "</td>" & vbCrLf
            if isRequired then 
                .Write "<td style=""color: Red""> &nbsp; *</td>"  & vbCrLf
            else
                .Write "<td/>" & vbCrLf
            end if
            if hasAddtionalInfo then
                .Write "<td>" & vbCrLf
                .Write "<img src=""../content/images/info.gif"" class=""mouseover"" alt="""" title=""additional information"" border=""0"" data-info=""" & _
                        NullableReplace(CurrentField.AdditionalInformation, """", "&quot;") & """ onclick=""showAdditionalInfo(this)"" />" & vbCrLf
                .Write "</td>" & vbCrLf
            else
                .Write "<td/>" & vbCrLf
            end if
        end with
    end sub 

    sub BuildCustomControlTextArea(CurrentField)
        dim fieldName : fieldName = "CustomField" & CurrentField.RegistrationCustomFieldId
        dim hasAddtionalInfo : hasAddtionalInfo = cbool(len(CurrentField.AdditionalInformation) > 0)
        dim hasNotes : hasNotes = cbool(len(CurrentField.Notes) > 0)
        dim isRequired : isRequired = CurrentField.Required
        dim validatorName : validatorName = "val" & fieldName

        with response
            .Write "<td style=""text-align: right"">" & CurrentField.Name & ":</td>" & vbCrLf
            .Write "<td/>" & vbCrLf
            if isRequired then 
                .Write "<td style=""color: Red""> &nbsp; *</td>"  & vbCrLf
                if hasAddtionalInfo then
                    .Write "<td>" & vbCrLf
                    .Write "<img src=""../content/images/info.gif"" class=""mouseover"" alt="""" title=""additional information"" border=""0"" data-info=""" & _
                            NullableReplace(CurrentField.AdditionalInformation, """", "&quot;") & """ onclick=""showAdditionalInfo(this)"" />" & vbCrLf
                    .Write "</td>" & vbCrLf
                else
                    .Write "<td/>" & vbCrLf
                end if
            else
                if hasAddtionalInfo then
                    .Write "<td/>"
                    .Write "<td>" & vbCrLf
                    .Write "<img src=""../content/images/info.gif"" class=""mouseover"" alt="""" title=""additional information"" border=""0"" data-info=""" & _
                            NullableReplace(CurrentField.AdditionalInformation, """", "&quot;") & """ onclick=""showAdditionalInfo(this)"" />" & vbCrLf
                    .Write "</td>" & vbCrLf
                else
                    .Write "<td colspan=""3""/>" & vbCrLf
                end if
            end if
            .Write "</tr>" & vbCrLf
            .Write "<tr>" & vbCrLf
            .Write "<td colspan=""4"">" & vbCrLf

            .Write "<textarea id=""" & FieldName & """ name=""" & fieldName & """ style=""resize: none; width: 98%"" "

            .Write "class="""
            if hasNotes then .Write " tooltip"
            if isRequired then .Write " requiredCustom"
            .Write """ "

            if hasNotes then .write "data-powertip=""" & nullablereplace(CurrentField.Notes, """", "&quot;") & """ "
            if isRequired then .Write "data-reqValidatorId=""" & validatorName & """"

            .Write "></textarea>" & vbCrLf
            
            if isRequired then .Write "<span id=""" & validatorName & """ style=""color: red; display: none"">Required</span>" & vbCrLf

            .Write "</td>" & vbCrLf
        end with
    end sub

    sub BuildCustomControlTextField(CurrentField)
        dim fieldName : fieldName = "CustomField" & CurrentField.RegistrationCustomFieldId
        dim hasAddtionalInfo : hasAddtionalInfo = cbool(len(CurrentField.AdditionalInformation) > 0)
        dim hasNotes : hasNotes = cbool(len(CurrentField.Notes) > 0)
        dim isRequired : isRequired = CurrentField.Required
        dim validatorName : validatorName = "val" & fieldName

        with response
            .Write "<td style=""text-align: right"">" & CurrentField.Name & ":</td>" & vbCrLf
            .Write "<td>" & vbCrLf
            .Write "<input type=""text"" name=""" & fieldName & """ id=""" & fieldName & """ "

            .Write "class="""
            if hasNotes then .Write " tooltip"
            if isRequired then .Write " requiredCustom"
            .Write """ "

            if hasNotes then .write "data-powertip=""" & nullablereplace(CurrentField.Notes, """", "&quot;") & """ "
            if isRequired then .Write "data-reqValidatorId=""" & validatorName & """"

            .Write " value=""""/>" & vbCrLF

            if isRequired then .Write "<span id=""" & validatorName & """ style=""color: red; display: none"">Required</span>" & vbCrLf

            .Write "</td>" & vbCrLf
            if isRequired then 
                .Write "<td style=""color: Red""> &nbsp; *</td>"  & vbCrLf
            else
                .Write "<td/>" & vbCrLf
            end if
            if hasAddtionalInfo then
                .Write "<td>" & vbCrLf
                .Write "<img src=""../content/images/info.gif"" class=""mouseover"" alt="""" title=""additional information"" border=""0"" data-info=""" & _
                        NullableReplace(CurrentField.AdditionalInformation, """", "&quot;") & """ onclick=""showAdditionalInfo(this)"" />" & vbCrLf
                .Write "</td>" & vbCrLf
            else
                .Write "<td/>" & vbCrLf
            end if
        end with
    end sub

    sub BuildTimeOptions(OpenTime, CloseTime)
        dim currentTime, endTime
        currentTime = OpenTime
        endTime = FormatDateTime(DateAdd("h", -2, GetTimeString(CloseTime)), vbShortTime)

        while currentTime <= endTime
            dim myEvents
            set myEvents = new ScheduledEventCollection

            myEvents.LoadByDateAndTime Request.QueryString("dt"), currentTime

            if myEvents.TotalPatrons < CInt(Settings(SETTING_REGISTRATION_MAXPLAYERS)) then
                response.Write "<option value='" & currentTime & "'>" & GetTimeString(currentTime) & "</option>"
            end if

            currentTime = FormatDateTime(DateAdd("n", CInt(Settings(SETTING_BOOKING_INTERVAL)), GetTimeString(currentTime)), vbShortTime)
        wend
    end sub

    function GetCancelEditText()
        dim text
    
        if AllowsCancelOrEdit() then
            if cbool(Settings(SETTING_REGISTRATION_ALLOWCANCELLATION)) and cbool(Settings(SETTING_REGISTRATION_ALLOWEDIT)) then
                text = "I understand that this deposit is non refundable. Cancellations must be done a minimum of " & Settings(SETTING_REGISTRATION_MAXDAYS_CANCEL) & _
                       " days in advance and any changes to the time or date must be done a minimum of " & Settings(SETTING_REGISTRATION_MAXDAYS_EDIT) & _
                       " days in advance, otherwise the entire deposit will be forfeited."
            elseif cbool(Settings(SETTING_REGISTRATION_ALLOWCANCELLATION)) then
                text = "I understand that this deposit is non refundable, and cancellations must be done a minimum of " & Settings(SETTING_REGISTRATION_MAXDAYS_CANCEL) & _
                       " days in advance, otherwise the entire deposit will be forfeited."
            elseif cbool(Settings(SETTING_REGISTRATION_ALLOWEDIT)) then
                text = "I understand that this deposit is non refundable, and any changes to the time or date must be done a minimum of " & _
                       Settings(SETTING_REGISTRATION_MAXDAYS_EDIT) & " days in advance, otherwise the entire deposit will be forfeited."
            end if
        end if
        
        GetCancelEditText = text      
    end function

    function GetPaymentText(PaymentTypeId, Amount)
        dim text
        text = "$" 
        text = text & FormatNumber(cdbl(Amount), 2)

        select case cstr(PaymentTypeId)
            case PAYMENTTYPE_BYEVENT
                'text = text & "event"
            case PAYMENTTYPE_BYPLAYER
                text = text & " / person"
        end select
        
        GetPaymentText = text
    end function
%>