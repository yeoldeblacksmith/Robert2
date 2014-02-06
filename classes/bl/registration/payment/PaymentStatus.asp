<%
class PaymentStatus
    'attributes
    private mn_Id, ms_Desc

    ' methods
    public sub LoadById(StatusId)
        Id = StatusId

        select case StatusId
            case PAYMENTSTATUS_UNPAID
                Description = "Unpaid"            
            case PAYMENTSTATUS_CONFLICTED
                Description = "Conflicted"
            case PAYMENTSTATUS_CANCELLED
                Description = "Cancelled"
            case PAYMENTSTATUS_FAILED
                Description = "Payment Failed"
            case PAYMENTSTATUS_PAID
                Description = "Paid"
            case PAYMENTSTATUS_PAIDCONFLICTED
                Description = "Paid - Conflicted"
            case PAYMENTSTATUS_REFUNDED
                Description = "Refunded"
            case PAYMENTSTATUS_CHANGEDPAYABLE
                Description = "Changed - Payable"
            case PAYMENTSTATUS_CHANGEDRECEIVABLE
                Description = "Changed - Receivable"
        end select
    end sub

    ' properties
    public property get Id
        Id = mn_Id
    end property

    public property let Id(value)
        mn_Id = value
    end property

    public property get Description
        Description = ms_Desc
    end property

    public property let Description(value)
        ms_Desc = value
    end property
end class
%>