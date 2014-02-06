<%
class OrderDetail
    'attributes
    private mn_EventId, mn_LineNumber, mn_ItemNumber, _
            ms_ItemDescription, mn_Quantity, md_Price, _
            mdt_CreateDateTime, ms_PaymentTransaction, mdt_PaymentDateTime, _
            md_PaymentAmount, mb_IsAdjustment

    private mn_OrgLineNumber, mn_OrgItemNumber, ms_OrgItemDescription, _
            mn_OrgQuantity, md_OrgPrice, mdt_OrgCreateDateTime, _
            ms_OrgPaymentTransaction, mdt_OrgPaymentDateTime, md_OrgPaymentAmount, _
            mb_OrgIsAdjustment

    'ctor
    public sub Class_Initialize()
        EventId = 0
        LineNumber = 0
        ItemNumber = 0
        Quantity = 0
        Price = 0.0
        ItemDescription = ""
        PaymentTransaction = ""
        PaymentAmount = 0.0
        IsAdjustment = false
    end sub

    'methods
    public sub CopyOrignialValues()
        mn_OrgLineNumber = mn_LineNumber
        mn_OrgItemNumber = mn_ItemNumber
        ms_OrgItemDescription = ms_ItemDescription
        mn_OrgQuantity = mn_Quantity
        md_OrgPrice = md_Price
        mdt_OrgCreateDateTime = mdt_CreateDateTime
        ms_OrgPaymentTransaction = ms_PaymentTransaction
        mdt_OrgPaymentDateTime = mdt_PaymentDateTime
        md_OrgPaymentAmount = md_PaymentAmount
        mb_OrgIsAdjustment = mb_IsAdjustment
    end sub

    public function HasChanged()
        dim status : status = false

        if mn_OrgLineNumber <> mn_LineNumber then status = true
        if mn_OrgItemNumber <> mn_ItemNumber then status = true
        if ms_OrgItemDescription <> ms_ItemDescription then status = true
        if cint(mn_OrgQuantity) <> cint(mn_Quantity) then status = true
        if cdbl(md_OrgPrice) <> cdbl(md_Price) then status = true
        if mdt_OrgCreateDateTime <> mdt_CreateDateTime then status = true
        if ms_OrgPaymentTransaction <> ms_PaymentTransaction then status = true
        if mdt_OrgPaymentDateTime <> mdt_PaymentDateTime then status = true
        if cdbl(md_OrgPaymentAmount) <> cdbl(md_PaymentAmount) then status = true
        if mb_OrgIsAdjustment <> mb_IsAdjustment then status = true

        HasChanged = status
    end function

    public sub Load()
        LoadByEventIdAndLineNumber EventId, LineNumber
    end sub

    public sub LoadByEventIdAndLineNumber(nEventId, nLineNumber)
        dim odCon : set odCon = new OrderDetailDataConnection
        dim results : results = odCon.GetOrderDetail(nEventId, nLineNumber)

        if UBound(results) > 0 then
            EventId = results(ORDERDETAIL_INDEX_EVENTID, 0)
            LineNumber = results(ORDERDETAIL_INDEX_LINENUMBER, 0)
            ItemNumber = results(ORDERDETAIL_INDEX_ITEMNUMBER, 0)
            ItemDescription = results(ORDERDETAIL_INDEX_ITEMDESCRIPTION, 0)
            Quantity = results(ORDERDETAIL_INDEX_QUANTITY, 0)
            Price = results(ORDERDETAIL_INDEX_PRICE, 0)
            CreateDateTime = results(ORDERDETAIL_INDEX_CREATEDATETIME, 0)
            PaymentTransaction = results(ORDERDETAIL_INDEX_PAYMENTTRANSACTION, 0)
            PaymentDateTime = results(ORDERDETAIL_INDEX_PAYMENTDATETIME, 0)
            PaymentAmount = results(ORDERDETAIL_INDEX_PAYMENTAMOUNT, 0)
            IsAdjustment = results(ORDERDETAIL_INDEX_ISADJUSTMENT, 0)
        end if

        CopyOrignialValues
    end sub

    public sub Save()
        dim odCon : set odCon = new OrderDetailDataConnection
        odCon.SaveOrderDetail EventId, LineNumber, ItemNumber, _
                              ItemDescription, Quantity, Price, _
                              CreateDateTime, PaymentTransaction, PaymentDateTime, _
                              PaymentAmount, IsAdjustment
    end sub

    'properties
    public property Get EventId
        EventId = mn_EventId
    end property

    public property Let EventId(value)
        mn_EventId = value
    end property
    
    public property Get LineNumber
        LineNumber = mn_LineNumber
    end property
    
    public property Let LineNumber(value)
        mn_LineNumber = value
    end property

    public property Get ItemNumber
        ItemNumber = mn_ItemNumber
    end property
    
    public property Let ItemNumber(value)
        mn_ItemNumber = value
    end property
        
    public property Get ItemDescription
        ItemDescription = ms_ItemDescription
    end property
    
    public property Let ItemDescription(value)
        ms_ItemDescription = value
    end property
    
    public property Get Quantity
        Quantity = mn_Quantity
    end property
    
    public property Let Quantity(value)
        mn_Quantity = cint(value)
    end property
    
    public property Get Price
        Price = md_Price
    end property
    
    public property Let Price(value)
        md_Price = cdbl(value)
    end property
    
    public property Get ExtendedPrice
        ExtendedPrice = cdbl(Price * Quantity)
    end property

    public property Get CreateDateTime
        CreateDateTime = mdt_CreateDateTime
    end property
    
    public property Let CreateDateTime(value)
        mdt_CreateDateTime = value
    end property
    
    public property Get PaymentTransaction
        PaymentTransaction = ms_PaymentTransaction
    end property
    
    public property Let PaymentTransaction(value)
        ms_PaymentTransaction = value
    end property
    
    public property Get PaymentDateTime
        PaymentDateTime = mdt_PaymentDateTime
    end property
    
    public property Let PaymentDateTime(value)
        mdt_PaymentDateTime = value
    end property
    
    public property Get PaymentAmount
        PaymentAmount = cdbl(md_PaymentAmount)
    end property
    
    public property Let PaymentAmount(value)
        md_PaymentAmount = value
    end property

    public property get IsPaid
        IsPaid = len(PaymentDateTime) > 0 and cbool(PaymentAmount >= ExtendedPrice)
    end property

    public property get IsAdjustment
        IsAdjustment = mb_IsAdjustment
    end property

    public property let IsAdjustment(value)
        mb_IsAdjustment = value
    end property
end class
%>