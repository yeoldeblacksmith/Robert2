<%
class ScheduledEventPaymentCollection
    ' attributes
    private mo_List

    'ctor
     public sub Class_Initialize
        set mo_List = new ArrayList
    end sub

    'methods
    public sub Add(value)
        mo_List.Add value
    end sub

    public sub LoadByEvent(EventId)
        dim myCon
        set myCon = new ScheduledEventPaymentDataConnection

        dim payments
        payments = myCon.GetAllScheduledEventPaymentByEventId(EventId)

        if UBound(payments) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(payments, 2) 
                dim myPayment
                set myPayment = new ScheduledEventPayment

                myPayment.EventId = payments(SCHEDULEDEVENTPAYMENT_INDEX_EVENTID, i)
                myPayment.StatusDateTime = payments(SCHEDULEDEVENTPAYMENT_INDEX_STATUSDATETIME, i)
                myPayment.PaymentAmount = payments(SCHEDULEDEVENTPAYMENT_INDEX_PAYMENTAMOUNT, i)
                myPayment.PaymentStatus = payments(SCHEDULEDEVENTPAYMENT_INDEX_PAYMENTSTATUS, i)
                myPayment.PayPalTransactionId = payments(SCHEDULEDEVENTPAYMENT_INDEX_PAYPALTRANSACTIONID, i)

                mo_List.Add myPayment
            next
        end if
    end sub

    public sub Remove(value)
        mo_List.Remove value
    end sub

    'proeprties
    public property get Count
        Count = mo_List.Count
    end property

    Public Default Property Get Item(index)
        set Item = mo_List(index)
    end property
end class
%>