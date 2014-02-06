<!--#include file="../classes/IncludeList.asp"-->
<%
    const AJAXACTION_PAYMENT_GETLASTPAYMENTSTATUS = "1"

    select case Request.QueryString(QUERYSTRING_VAR_ACTION)
        case AJAXACTION_PAYMENT_GETLASTPAYMENTSTATUS
            WriteLastPaymentStatus request.QueryString(QUERYSTRING_VAR_ID)
    end select

    private sub WriteLastPaymentStatus(EncodedEventId)
        dim payment : set payment = new ScheduledEventPayment
        payment.LoadLastPaymentByEventId DecodeId(EncodedEventId)

        response.Write lcase(payment.PaymentStatus)
    end sub
%>