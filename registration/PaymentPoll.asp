<!--#include file="../classes/IncludeList.asp"-->
<%
    if len(Request.Form(QUERYSTRING_VAR_ID)) = 0 then response.Redirect("ErrorPage.asp")

    dim myEvent : set myEvent = new ScheduledEvent
    myEvent.Load DecodeId(Request.Form(QUERYSTRING_VAR_ID))

    dim expireDateTime : expireDateTime = dateadd("n", 10, myEvent.CreateDateTime)
    dim totalRemainingSeconds : totalRemainingSeconds = datediff("s", now(), expireDateTime)
    dim remainingMinutes : remainingMinutes = 0
    dim remainingSeconds : remainingSeconds = 0

    if totalRemainingSeconds > 0 then
        remainingMinutes = totalRemainingSeconds \ 60
        remainingSeconds = totalRemainingSeconds mod 60
    end if
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Waiting for payment</title>
        <meta name="viewport" content="width=device-width, initial-scale=1"/>
        <meta http-equiv="cache-control" content="max-age=0" /> 
        <meta http-equiv="cache-control" content="no-cache" /> 
        <meta http-equiv="expires" content="0" /> 
        <meta http-equiv="expires" content="Tue, 01 Jan 1980 1:00:00 GMT" /> 
        <meta http-equiv="pragma" content="no-cache" />
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>

        <link type="text/css" rel="stylesheet" href="../content/css/jquery.mobile-1.1.0.css" />
        <link type="text/css" rel="Stylesheet" href="../content/css/eventregistration.css?<%= ANTI_CACHE_STRING %>" />
        <link type="text/css" rel="Stylesheet" href="../content/css/waiver.css?<%= ANTI_CACHE_STRING %>" />

        <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.mobile-1.1.0.min.js"></script>
        <script type="text/javascript">
            var minutes = <%= remainingMinutes %>;
            var seconds = <%= remainingSeconds %>;
            var paymentCompleted = false;

            $(document).ready(function () {
                displayTime();
                runCounter();
            });

            function displayTime() {
                var secondString;

                if (seconds >= 10) {
                    secondString = seconds;
                } else {
                    secondString = "0" + seconds;
                }

                $("#RemainingTime").html(minutes + ":" + secondString);
            }

            function pollPayment() {
                $.ajax({
                    url: '../ajax/payment.asp?act=1&id=<%= Request.Form(QUERYSTRING_VAR_ID) %>',
                    error: function () { alert("problems encountered polling for payments"); },
                    success: function (results) {
                        if (results == "completed") {
                            paymentCompleted = true;
                        }
                    }
                });
            }

            function runCounter() {
                if (paymentCompleted == false) {
                    if (seconds <= 0) {
                        seconds = 60;
                        minutes -= 1;
                    }

                    if (minutes <= -1) {
                        seconds = 0;
                        minutes += 1;

                        $("#divWaiting").css("display", "none");
                        $("#divForm").css("display", "none");
                        $("#divExpired").css("display", "block");
                    } else {
                        seconds -= 1;
                        displayTime();

                        if (seconds % 5 == 0) { pollPayment(); }

                        setTimeout("runCounter()", 1000);
                    }
                } else {
                    document.forms[0].action = "PaymentReceived.asp";
                    document.forms[0].target = "_self";
                    document.forms[0].submit();
                }
            }
        </script>
    </head>
    <body>
        <div class="waiverSection blockCenter" style="width: 550px; padding: 15px">
            <img src="<%= LOGO_URL %>" style="display: block; margin: auto" />
            <h3 class="center">Waiting for payment</h3>

            <div id="divWaiting">
                <p>We are awaiting payment for your reservation. We can only reserve the time for you for 10 minutes. If we have not received a payment by then,
                    we cannot gaurantee that you will be able to get that time. If we have not received payment, and another party is booked for that time, we
                    will notify you at the email address supplied.
                </p>

                <p>
                    <b>Time Reserved for: </b>
                    <span id="RemainingTime" />
                </p>
            </div>

            <div id="divExpired" style="display: none;">
                <p>Unfortunately, the time has expired and we can no longer hold this time for you. You can still make a payment. If payment is received before
                    someone else claims the spot, there is not a problem. If someone else claims the spot before you can make the payment, an email will be sent
                    to the email address you provided.
                </p>
            </div>

            <div id="divPaymentInfo">
                <table id="tblPaymentDetails" class="blockCenter">
                    <tr>
                        <th>Description</th>
                        <th>Qty</th>
                        <th>Price</th>
                        <th>Ext. Price</th>
                    </tr>
<%
    dim details : set details = new OrderDetailCollection
    details.LoadByEventId myEvent.EventId

    for i = 0 to details.Count - 1
        with response
            .Write "<tr>" & vbCrLf
            .Write "<td class=""inlineRight"">" & vbCrLf
            .Write details(i).ItemDescription
            .Write "</td>" & vbCrLf
            .Write "<td class=""inlineRight"">" & vbCrLf
            .Write details(i).Quantity
            .Write "</td>" & vbCrLf
            .Write "<td class=""inlineRight"">" & vbCrLf
            .Write FormatNumber(details(i).Price,2)
            .Write "</td>" & vbCrLf
            .Write "<td class=""inlineRight"">" & vbCrLf
            .Write FormatNumber(details(i).ExtendedPrice, 2)
            .Write "</td>" & vbCrLf
            .Write "</tr>" & vbCrLf
        end with
    next
%>
                    <tr>
                        <td colspan="4">
                            <hr/>
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3" />
                        <td class="inlineRight">
                            <%= FormatNumber(details.TotalAmountOwed, 2) %>
                        </td>
                    </tr>
                </table>
            </div>

            <div id="divForm">
                <p>Please use the button below to make your payment.</p>
                <form action="<%= PAYPAL_IPN_ADDRESS %>" method="post" target="_blank">
                    <input type="hidden" name="<%= QUERYSTRING_VAR_ID %>" value="<%= Request.Form(QUERYSTRING_VAR_ID) %>" />
	                <input type="hidden" name="charset" value="utf-8"> 
	                <input type="hidden" name="cmd" value="_cart" />
                    <input type="hidden" name="upload" value="1" />
                    <input type="hidden" name="<%= PAYPAL_VARIABLE_EVENTID %>" value="<%= myEvent.EventId %>" />
                    <input type="hidden" name="return" value="<%= SiteInfo.VantoraUrl %>/registration/confirmation.asp" />
	                <input type="hidden" name="cancel_return" value="<%= SiteInfo.VantoraUrl %>/registration/paymentcancelled.asp?id=<%= EncodeId(myEvent.EventId) %>" />
                    <input type="hidden" name="notify_url" value="<%= SiteInfo.VantoraUrl %>/registration/paypalipnlistener.asp" />
	                <input type="hidden" name="business" value="<%= Settings(SETTING_PAYMENT_PAYPALADDRESS) %>" />
	                <input type="hidden" name="no_note" value="0" />
	                <input type="hidden" name="cbt" value="Return to <%= SiteInfo.Name %>" />
                    <input type="hidden" name="currency_code" value="<%= GetCurrencyValue() %>" />
                    <input type="hidden" name="rm" value="2" />

    <%
                for j = 0 to details.Count - 1
                    response.Write "<input type=""hidden"" name=""" & PAYPAL_VARIABLE_ITEMDESCRIPTION & "_" & j + 1 & """ value=""" & details(j).ItemDescription & """/>" & vbCrLf
                    response.Write "<input type=""hidden"" name=""" & PAYPAL_VARIABLE_ITEMLINENUMBER & "_" & j + 1 & """ value=""" & details(j).LineNumber & """/>" & vbCrLf
                    response.Write "<input type=""hidden"" name=""" & PAYPAL_VARIABLE_ITEMPRICE & "_" & j + 1 & """ value=""" & FormatNumber(details(j).Price, 2) & """/>" & vbCrLf
                    response.Write "<input type=""hidden"" name=""" & PAYPAL_VARIABLE_ITEMQTY & "_" & j + 1 & """ value=""" & details(j).Quantity & """/>" & vbCrLf
                next
    %>

                    <img src="../content/images/paypal-paynow.gif" style="display: block; margin: auto; cursor: pointer" onclick="javascript: document.forms[0].submit();" />
                </form>
            </div>
        </div>
    </body>
</html>
<%
    function GetCurrencyValue()
        select case SiteInfo.Country 
            case "CA"
                GetCurrencyValue = "CAD"
            case "US"
                GetCurrencyValue = "USD"
        end select
    end function
%>