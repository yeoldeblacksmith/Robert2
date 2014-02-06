<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file="../common/EventIncludeList.asp" -->
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Edit Registration</title>
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
            $(document).live('pageinit', function () {
                $.ajax({
                    url: '../common/eventajax.asp?act=5&id=' + escape($("input[name=EventId]").val()),
                    success: function (data) { $("#form").html(data); },
                    error: function () { alert("problems encountered"); }
                });
            });

            function cancelRegistration() {
                if (confirm("Are you sure?")) {
                    var url = '../common/eventajax.asp?act=4&id=' + escape($("input[name=EventId]").val()) + '&de=true';

                    $.ajax({
                        url: url,
                        success: function () {
                            redirectToHome();
                        },
                        error: function () { alert('problems encountered'); }
                    });
                }
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

            function redirectToHome() {
                window.location = "cancelconfirmation.asp";
            }

            function reloadTime(selectedDate) {
                var timeList = $("select[name=Time]");

                timeList.empty();

                $.ajax({
                    url: '../common/availabilityajax.asp?av=2&dt=' + selectedDate,
                    success: function (data) { timeList.html(data); },
                    error: function () { alert("problems encountered"); }
                });

                setOpenHours(selectedDate);

            }

            function saveRegistration() {
                if (validateForm()) {
                    var url = '../common/eventajax.asp?act=3&id=' + $("input[name=EventId]").val() +
                                                    '&dt=' + escape($("input[name=Date]").val()) +
                                                    '&st=' + escape($("select[name=Time]").val()) +
                                                    '&np=' + escape($("input[name=NumberOfGuests]").val()) +
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

            function setOpenHours(selectedDate) {
                $.ajax({
                    url: '../common/availabilityajax.asp?av=3&dt=' + selectedDate,
                    success: function (data) { $("span[name=openHours]").html(data); },
                    error: function () { alert("problems encountered"); }
                });
            }

            function setSelectedDay(selectedDate) {
                $("input[name=Date]").val(selectedDate);
            }

            function submitVerification() {
                if (validateVerification()) {
                    $.ajax({
                        url: '../common/eventajax.asp?act=6&id=' + escape($("input[name=EventId]").val()) +
                                                       '&em=' + escape($("input[name=Email]").val()),
                        success: function (data) { $("#form").html(data); },
                        error: function () { alert("problems encountered"); }
                    });
                } else {
                    $("input[name=Email]").focus();
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
                <input type="hidden" name="EventId" value="<%= Request.QueryString(QUERYSTRING_VAR_ID) %>"/>
            
                <span id="form"></span>
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
