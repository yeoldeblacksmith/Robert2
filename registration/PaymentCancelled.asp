<!--#include file="../classes/IncludeList.asp" -->
<%
    dim myEvent : set myEvent = new ScheduledEvent
    dim details : set details = new OrderDetailCollection

    if Request.QueryString(QUERYSTRING_VAR_ID) = 0 then
        response.Redirect("ErrorPage.asp")
    else
        myEvent.Load DecodeId(Request.QueryString(QUERYSTRING_VAR_ID))
        details.LoadByEventId myEvent.EventId
    end if
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Event Registration Payment Cancelled</title>
        <meta name="viewport" content="width=device-width, initial-scale=1"/>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>

        <link type="text/css" rel="Stylesheet" href="../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../content/css/waiver.css" />

        <script type="text/javascript" src="../content/js/common.js"></script>
        <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
    </head>

    <body>
        <div class="waiverSection blockCenter" style="width: 550px">
            <img src="<%= LOGO_URL %>" style="display: block; margin: auto" />
            <h3 style="text-align: center">Payment Cancelled</h3>
    
            <blockquote>
                <p>
                    Please remember that we only keep that time available for only 10 minutes. To make sure that someone else does not
                    get the time you wanted, please pay the deposit as soon as possible.
                </p>
            
                <p>Please use the button below to make your payment.</p>
            </blockquote>

            <form action="<%= PAYPAL_IPN_ADDRESS %>" method="post">
                <input type="hidden" name="<%= QUERYSTRING_VAR_ID %>" value="<%= myEvent.EventId %>" />
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
        <br /><br /><br />
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