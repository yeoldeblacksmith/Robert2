<%
class ScheduledEventPayment
    ' attributes
    private mn_EventId, mdt_StatusDateTime, md_PaymentAmount, _
            ms_PaymentStatus, ms_PayPalTransactionId

    ' methods
    public function IsPaid()
        IsPaid = cbool(StrComp(PaymentStatus, PAYPAL_PAYMENTSTATUS_PAID, 1) = 0)
    end function

    public sub LoadLastPaymentByEventId(nEventId)
        EventId = nEventId

        LoadLastPayment
    end sub

    public sub LoadLastPayment()
        dim myCon : set myCon = new ScheduledEventPaymentDataConnection
        dim results : results = myCon.GetLastScheduledEventPaymentByEventId(EventId)

        if UBound(results) > 0 then
            EventId = results(SCHEDULEDEVENTPAYMENT_INDEX_EVENTID, 0)
            StatusDateTime = results(SCHEDULEDEVENTPAYMENT_INDEX_STATUSDATETIME, 0)
            PaymentAmount = results(SCHEDULEDEVENTPAYMENT_INDEX_PAYMENTAMOUNT, 0)
            PaymentStatus = results(SCHEDULEDEVENTPAYMENT_INDEX_PAYMENTSTATUS, 0)
            PayPalTransactionId = results(SCHEDULEDEVENTPAYMENT_INDEX_PAYPALTRANSACTIONID, 0)
        end if
    end  sub

    public sub Save()
        dim myCon : set myCon = new ScheduledEventPaymentDataConnection
        myCon.SaveScheduledEventPayment EventId, StatusDateTime, PaymentAmount, PaymentStatus, PayPalTransactionId
    end sub

    ' properties
    public property get EventId    
        EventId = mn_EventId
    end property

    public property let EventId(value)
        mn_EventId = value
    end property

    public property get StatusDateTime
        StatusDateTime = mdt_StatusDateTime
    end property

    public property let StatusDateTime(value)
        mdt_StatusDateTime = value
    end property

    public property get PaymentAmount
        PaymentAmount = md_PaymentAmount
    end property

    public property let PaymentAmount(value)
        md_PaymentAmount = value
    end property

    public property get PaymentStatus
        PaymentStatus = ms_PaymentStatus
    end property

    public property let PaymentStatus(value)
        ms_PaymentStatus = value
    end property

    public property get PayPalTransactionId
        PayPalTransactionId = ms_PayPalTransactionId
    end property

    public property let PayPalTransactionId(value)
        ms_PayPalTransactionId = value
    end property
end class
%>