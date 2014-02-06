<!--#include file="../classes/IncludeList.asp"-->
<%
    if len(Request.QueryString(QUERYSTRING_VAR_ID)) = 0 then response.Redirect("ErrorPage.asp")

    dim myEvent : set myEvent = new ScheduledEvent
    myEvent.Load DecodeId(Request.QueryString(QUERYSTRING_VAR_ID))

    dim bIsPaid : bIsPaid = cbool(myEvent.PaymentStatusId = PAYMENTSTATUS_PAID or _
                                  myEvent.PaymentStatusId = PAYMENTSTATUS_MARKEDPAID or _
                                  myEvent.PaymentStatusId = PAYMENTSTATUS_PAIDCONFLICTED) 

    dim details : set details = new OrderDetailCollection
    details.LoadByEventId myEvent.EventId
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title></title>
</head>
    <body>
        <% if bIsPaid then %>
        <form action="confirmation.asp" method="post">
            <input type="hidden" name="<%= PAYPAL_VARIABLE_EVENTID %>" value="<%= myEvent.EventId %>" />
        </form>
        <% else %>
        <form action="<%= PAYPAL_IPN_ADDRESS %>" method="post">
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

        </form>
        <% end if %>
        <script type="text/javascript">
            document.forms[0].submit();
        </script>
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