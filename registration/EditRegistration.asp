<!DOCTYPE html>
<!--#include file="../classes/IncludeList.asp" -->
<%
    'RedirectMobileToSpecificPage "mobile/editreg.asp?" & QUERYSTRING_VAR_ID & "=" & Request.QueryString(QUERYSTRING_VAR_ID)
    if cbool(settings(SETTING_MODULE_REGISTRATION)) = false then response.Redirect("disabled.asp")
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Edit Event Registration</title>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <meta name="viewport" content="width=device-width, maximum-scale=1.0, minimum-scale=1.0" /> 

        <link type="text/css" rel="stylesheet" href="../content/css/jquery.fancybox-1.3.4.css" />
        <link type="text/css" rel="stylesheet" href="../content/css/jquery.mobile-1.3.2.css" />
        <link type="text/css" rel="stylesheet" href="../content/css/jquery.powertip-light.css" />
        <link type="text/css" rel="Stylesheet" href="../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../content/css/waiver.css" />

        <script type="text/javascript" src="../content/js/common.js"></script>
        <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.mobile-1.3.2.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.fancybox-1.3.4.pack.js"></script>
        <script type="text/javascript" src="../content/js/jquery.powertip.min.js"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                $.ajax({
                    url: '../ajax/event.asp?act=5&id=' + escape($("input[name=EventId]").val()),
                    success: function (data) { $("#form").html(data).trigger("create"); },
                    error: function () { alert("problems encountered"); }
                });
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
                            addDetailRowToCart('Base Registration Fee', 1, baseFee, baseFee);
                            break;
                        case <%=PAYMENTTYPE_BYPLAYER%>:
                            if (players == 0) {
                                baseFee = 0;
                            } else {                               
                                var extPrice = players * baseFee;
                                addDetailRowToCart('Base Registration Fee', players, baseFee, extPrice);

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
                } else {
                    detailTable.css("display", "table");
                }
            }

            function cancelRegistration() {
                if (confirm("Are you sure?")) {
                    var url = '../ajax/event.asp?act=4&id=' + escape($("input[name=EventId]").val()) + 
                                                '&ss=<%=PAYMENTSTATUS_CANCELLED_USER%>&de=true&rs=true';

                    $.ajax({
                        url: url,
                        sync: true,
                        success: function () {
                            window.location.href = "cancelconfirmation.asp";
                        },
                        error: function () { alert('problems encountered'); }
                    });
                }

                return false;

            }

            function handleUserMaxLength() {
                var commentField = $("textarea[name=UserComments]");

                if (commentField.val().length <= commentField.attr("maxlength")) {
                    var remains = commentField.attr("maxlength") - commentField.val().length;
                    $("span[name=userCharCounter]").html("(" + remains + " characters available)");
                    return true;
                } else {
                    return false;
                }
            }

            function hideAvailableTimeWindow() {
                //$("#popupTimes").popup("close");
                $.fancybox.close();
            }

            //function initializeForm(EventTypeId, NumberOfGuests, StartTime) {
            function initializeForm() {
                //$("#NumberOfGuests").val(NumberOfGuests);
                setOpenHours($("#SelectedDate").val());
                calculateTotals();

                $("#btnCalendar").fancybox({
                    'type': 'iframe',
                    'href': 'MiniCalendar.asp',
                    'titlePosition': 'outside',
                    'title': '&copy;GatSplat',
                    'width': 420,
                    'height': 365,
                    'transitionIn': 'elastic',
                    'transitionOut': 'elastic',
                    'speedIn': 200,
                    'speedOut': 200
                });

                $(".tooltip").powerTip({ mouseOnToPopup: true, placement: 'n', smartPlacement: true });
                $("input:not([type='hidden']):not([type='checkbox'])").textinput({ preventFocusZoom: true });
                $("select").selectmenu({ preventFocusZoom: true });
            }

            function NumberOfGuests_onblur() {
                calculateTotals();

                if (!!validateNumberOfGuests() == false) {
                    $("#btnSubmit").button("disable");
                }
            }

            function redirectToHome() {
                window.location.href = '<%= SiteInfo.HomeUrl %>';
            }

            function resizeFancyBox(){
                $('#fancybox-content').height($('#fancybox-frame').contents().find('div.waiverSection').height() + 5);
                $.fancybox.center();
            }

            function setOpenHours(selectedDate) {
                $.ajax({
                    url: '../ajax/availability.asp?av=3&dt=' + selectedDate,
                    success: function (data) { $("span[name=openHours]").html(data); },
                    error: function () { alert("problems encountered"); }
                });
            }

            function setSelectedDay(selectedDate) {
                $("#SelectedDate").val(selectedDate);
            }

            function setStartTime(Time, Display) {
                $("#lblStartTime").html(Display);
                $("#Time").val(Time);

                if ($("#btnSubmit").length > 0) { $("#btnSubmit").button("enable"); }
            }

            function showAdditionalInfo(customField) {
                $.fancybox({
                    'content': $(customField).data("info"),
                    'showCloseButton': true
                });
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

            function submitVerification() {
                if (validateVerification()) {
                    $.ajax({
                        url: '../ajax/event.asp?act=6&id=' + escape($("input[name=EventId]").val()) +
                                                       '&em=' + escape($("input[name=Email]").val()),
                        success: function (data) { $("#form").html(data).trigger("create"); },
                        error: function () { alert("problems encountered"); }
                    });
                } else {
                    $("input[name=Email]").focus();
                }
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

                return valid;
            }

            function validateForm() {
                //return (validateEmail() & validatePhoneNumber() & validateContactName() & validateNumberOfGuests());

                // we are having to play games here
                var valid = !!(validateEmail() & validatePhoneNumber() & validateContactName() &
                               validateNumberOfGuests() & validateRequiredCustomFields());

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

            function validateVerification() {
                var valid = true;
                var email = $("input[name=Email]").val().toLowerCase();
                var emailRegExp = /^([a-zA-Z0-9_\.\-\+])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;

                if (email.length > 0) {
                    $("span[name=ValEmailRequired]").css("display", "none");

                    if (emailRegExp.test(email)) {
                        $("span[name=ValEmail]").css("display", "none");
                    } else {
                        $("span[name=ValEmail]").css("display", "inline");
                        valid = false;
                    } // regExp test
                } else {
                    $("span[name=ValEmailRequired]").css("display", "inline");
                    $("span[name=ValEmail]").css("display", "none");
                    valid = false;
                } // length test

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
        <form action="Save.asp" id="myForm" method="post" data-ajax="false">
            <input type="hidden" name="EventId" value="<%= Request.QueryString(QUERYSTRING_VAR_ID) %>"/>
            
            <div id="form" class="waiverSection blockCenter" style="width: 550px"></div>
        </form>
    </body>
</html>