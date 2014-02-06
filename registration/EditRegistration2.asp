<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="common/EventIncludeList.asp" -->
<%
    RedirectMobileToSpecificPage "mobile/editreg.asp?" & QUERYSTRING_VAR_ID & "=" & Request.QueryString(QUERYSTRING_VAR_ID)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Edit Event Registration</title>
        <meta name="description" content="Indoor Paintball fields perfect for birthday parties, bachelor parties as well as corporate team building events or any event in DFW." />
        <meta Name="keywords" Content="paintball games, dfw, venue, dallas, paintball field, indoor paintball, paintball supplies, gatsplat, birthday parties,  ">
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <!--#include file="../nghead.asp"-->

        <meta name="google-site-verification" content="IUE9L0Evw9Ovgc86mXY-AMjCjyFXteKGTp0NiB6eE7U" />
        <meta name="viewport" content="width=device-width, maximum-scale=1.0, minimum-scale=1.0" /> 

        <link type="text/css" rel="Stylesheet" href="content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="content/css/jquery.fancybox-1.3.4.css" />

        <script type="text/javascript" src="content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="content/js/jquery.fancybox-1.3.4.pack.js"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                $.ajax({
                    url: 'common/eventajax2.asp?act=5&id=' + escape($("input[name=EventId]").val()),
                    success: function (data) { $("#form").html(data); },
                    error: function () { alert("problems encountered"); }
                });
            });

            function addInvite(enc) {
                if (validateInvite()) {
                    var name = escape($("#txtGuestName").val());
                    var email = escape($("#txtGuestEmail").val());
                    var eventId = $("#eventId").val();
                    var guestCnt = $("#txtRsvpCount").val();
                    var status = $("#ddlRsvpStatus").val();
                    var sUrl = 'common/inviteajax.asp?act=1&id=' + eventId + "&nm=" + name + "&em=" + email + '&np=' + guestCnt + '&ss=' + status;

                    $.ajax({
                        url: sUrl,
                        error: function () { alert("problems encountered sending request"); },
                        success: function () { loadTable(); }
                    });
                }
            }

            function cancelRegistration() {
                if (confirm("Are you sure?")) {
                    var url = 'common/eventajax2.asp?act=4&id=' + escape($("input[name=EventId]").val()) + '&de=true&rs=true';

                    $.ajax({
                        url: url,
                        success: function () {
                            redirectToHome();
                        },
                        error: function () { alert('problems encountered'); }
                    });
                }
            }

            function deleteInvite(inviteId) {
                if (confirm("Are you sure?")) {
                    $.ajax({
                        url: 'common/inviteajax.asp?act=5&id=' + inviteId,
                        error: function () { alert("problems encountered deleting invitation"); },
                        success: function () { loadTable(); }
                    });
                }
            }

            function editInvite(inviteId) {
                $.ajax({
                    url: 'common/inviteajax.asp?act=26&id=' + $("#eventId").val() + '&nm=' + inviteId,
                    error: function () { alert("problems encountered"); },
                    success: function (data) {
                        var container = $("#form");

                        container.empty();
                        container.html(data);
                    }
                });
            }

            function handleMaxLength() {
                var commentField = $("textarea[name=UserComments]");

                if (commentField.val().length <= commentField.attr("maxlength")) {
                    var remains = commentField.attr("maxlength") - commentField.val().length;
                    $("span[name=userCharCounter]").html("(" + remains + " characters available)");
                    return true;
                } else {
                    return false;
                }
            }

            function handleMaxLengthInvitation() {
                var commentField = $("#UserComments");

                if (commentField.val().length <= commentField.attr("maxlength")) {
                    var remains = commentField.attr("maxlength") - commentField.val().length;
                    $("span[name=charcounter]").html("(" + remains + " characters available)");
                    return true;
                } else {
                    return false;
                }
            }

            function initializeForm(EventTypeId, NumberOfGuests, StartTime) {
                $("select[name=EventTypeId]").val(EventTypeId);
                $("select[name=NumberOfGuests]").val(NumberOfGuests);
                setOpenHours($("input[name=Date]").val());
                $("select[name=Time]").val(StartTime);

                $("button[name=btnCalendar]").fancybox({
                    'type': 'iframe',
                    'href': 'MiniCalendar.asp',
                    'titlePosition': 'outside',
                    'title': '&copy;GatSplat',
                    'width': 420,
                    'height': 325,
                    'transitionIn': 'elastic',
                    'transitionOut': 'elastic',
                    'speedIn': 200,
                    'speedOut': 200
                });
            }

            function loadTable() {
                showInviteForm();
            }

            function redirectToHome() {
                window.location = "cancelconfirmation.asp";
            }

            function reloadTime(selectedDate) {
                var timeList = $("select[name=Time]");

                timeList.empty();

                $.ajax({
                    url: 'common/availabilityajax.asp?av=2&dt=' + selectedDate,
                    success: function (data) { timeList.html(data); },
                    error: function () { alert("problems encountered"); }
                });

                setOpenHours(selectedDate);

            }

            function resendInvite(inviteId) {
                $.ajax({
                    url: 'common/inviteajax.asp?act=32&id=' + inviteId,
                    error: function () { alert("problems encountered"); },
                    success: function (data) { alert(data); }
                });
            }

            function saveRegistration() {
                if (validateForm()) {
                    var url = 'common/eventajax2.asp?act=3&id=' + $("input[name=EventId]").val() +
                                                    '&dt=' + escape($("input[name=Date]").val()) +
                                                    '&st=' + escape($("select[name=Time]").val()) +
                                                    '&np=' + escape($("select[name=NumberOfGuests]").val()) +
                                                    '&ap=' + escape($("input[name=AgeOfGuests]").val()) +
                                                    '&em=' + escape($("input[name=Email]").val()) +
                                                    '&ph=' + escape($("input[name=Phone]").val()) +
                                                    '&nm=' + escape($("input[name=Name]").val()) +
                                                    '&ev=' + escape($("select[name=EventTypeId]").val()) +
                                                    '&pn=' + escape($("input[name=PartyName]").val()) +
                                                    '&uc=' + escape($("textarea[name=UserComments]").val()) +
                                                    '&ac=' + escape($("input[name=AdminComments]").val()) +
                                                    '&rs=true&de=true';

                    $.ajax({
                        url: url,
                        success: function () {
                            window.location = "UpdateConfirmation.asp";
                        },
                        error: function () { alert("problems encountered"); }
                    });
                }
            }

            function sendInvites() {
                $.ajax({
                    url: 'common/inviteajax.asp?act=31&id=' + $("#eventId").val(),
                    error: function () { alert("problems encountered"); },
                    success: function (data) { alert(data); }
                });
            }

            function setOpenHours(selectedDate) {
                $.ajax({
                    url: 'common/availabilityajax.asp?av=3&dt=' + selectedDate,
                    success: function (data) { $("span[name=openHours]").html(data); },
                    error: function () { alert("problems encountered"); }
                });
            }

            function setSelectedDay(selectedDate) {
                $("input[name=Date]").val(selectedDate);
            }

            function showInviteForm() {
                $.ajax({
                    url: 'common/inviteajax.asp?act=25&id=' + $("#eventId").val(),
                    error: function () { alert("problems encountered"); },
                    success: function (data) { $("#form").html(data); }
                });
            }

            function showRegistrationForm() {
                $.ajax({
                    url: 'common/eventajax2.asp?act=10&id=' + $("#eventId").val(),
                    error: function () { alert("problems encountered"); },
                    success: function (data) { $("#form").html(data); }
                });
            }

            function submitVerification() {
                if (validateVerification()) {
                    $.ajax({
                        url: 'common/eventajax2.asp?act=6&id=' + escape($("input[name=EventId]").val()) +
                                                       '&em=' + escape($("input[name=Email]").val()),
                        success: function (data) { $("#form").html(data); },
                        error: function () { alert("problems encountered"); }
                    });
                } else {
                    $("input[name=Email]").focus();
                }
            }

            function updateInvite(inviteId) {
                if (validateInvite()) {
                    var name = escape($("#txtGuestName").val());
                    var email = escape($("#txtGuestEmail").val());
                    var guestCnt = $("#txtRsvpCount").val();
                    var status = $("#ddlRsvpStatus").val();
                    var comments = escape($("#UserComments").val());

                    $.ajax({
                        url: 'common/inviteajax.asp?act=11&id=' + inviteId + "&nm=" + name + "&em=" + email + '&np=' + guestCnt + '&ss=' + status + '&uc=' + comments,
                        error: function () { alert("problems encountered sending request"); },
                        success: function () { loadTable(); }
                    });
                }
            }

            function validateForm() {
                var valid = true;

                // validate contact email address
                var email = $("input[name=Email]").val().toLowerCase();
                var confirm = $("input[name=ConfirmEmail]").val().toLowerCase();
                var emailRegExp = /^([a-zA-Z0-9_\.\-\+])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
                var phoneRegExp = /^\(?(\d{3})\)?[- ]?(\d{3})[- ]?(\d{4})$/;

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

                // validate contact name
                var name = $("input[name=Name]").val();

                if (name.length == 0) {
                    $("span[name=ValName]").css("display", "inline");
                    valid = false;
                } else {
                    $("span[name=ValName]").css("display", "none");
                }

                return valid;
            }

            function validateInvite() {
                var valid = true;

                // validate contact email address
                var email = $("#txtGuestEmail").val().toLowerCase();
                var emailRegExp = /^([a-zA-Z0-9_\.\-\+])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;

                if (email.trim().length > 0) {
                    if (emailRegExp.test(email)) {
                        $("span[name=ValEmail]").css("display", "none");
                        $("span[name=ValEmailRequired]").css("display", "none");
                    } else {
                        $("span[name=ValEmail]").css("display", "inline");
                        valid = false;
                    } // regExp test
                } else {
                    $("span[name=ValEmailRequired]").css("display", "inline");
                    valid = false;
                } // length test

                // validate contact name
                var name = $("#txtGuestName").val();

                if (name.trim().length == 0) {
                    $("span[name=ValName]").css("display", "inline");
                    valid = false;
                } else {
                    $("span[name=ValName]").css("display", "none");
                }

                // validate rsvp count
                var guestCnt = $("#txtRsvpCount").val();

                if (guestCnt.length == 0) {
                    $("span[name=ValRsvpCount]").css("display", "none");
                } else {
                    if (isNaN(parseInt(guestCnt))) {
                        $("span[name=ValRsvpCount]").css("display", "inline");
                        valid = false;
                    } else {
                        if (guestCnt <= 0) {
                            $("span[name=ValRsvpCount]").css("display", "inline");
                            valid = false;
                        } else {
                            $("span[name=ValRsvpCount]").css("display", "none");
                        }
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
        </script>    
    </head>
    <body>
        <!--#include file="..\gatmenuhead.asp"--></blockquote>

        <input type="hidden" id="eventId" name="EventId" value="<%= Request.QueryString(QUERYSTRING_VAR_ID) %>"/>
            
        <span id="form"></span>

    </body>
</html>