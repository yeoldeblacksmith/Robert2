<%
class PaymentTypeCollection
    'attributes 
    private mo_List

    ' ctor
    public sub Class_Initialize()
        set mo_List = new ArrayList
    end sub

    'methods
    public sub Load
        dim paymentTypeNone, _
            paymentTypeByEvent, _
            paymentTypeByPlayer

        set paymentTypeNone = new PaymentType
        set paymentTypeByEvent = new PaymentType
        set paymentTypeByPlayer = new PaymentType

        paymentTypeNone.LoadById PAYMENTTYPE_NONE
        paymentTypeByEvent.LoadById PAYMENTTYPE_BYEVENT
        paymentTypeByPlayer.LoadById PAYMENTTYPE_BYPLAYER

        mo_List.Add paymentTypeNone
        mo_List.Add paymentTypeByEvent
        mo_List.Add paymentTypeByPlayer
    end sub

    'properties
    public property get Count
        Count = mo_List.Count
    end property

    Public Default Property Get Item(index)
        set Item = mo_List(index)
    end property
end class
%>