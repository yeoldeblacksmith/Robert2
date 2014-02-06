<%
class OrderDetailCollection
    'attributes 
    private mo_List

    ' ctor
    public sub Class_Initialize()
        set mo_List = new ArrayList
    end sub

    'methods
    public sub Add(value)
        mo_List.Add value
    end sub

    public sub DeleteByEventId(nEventId)
        dim odCon : set odCon = new OrderDetailDataConnection
        odCon.DeleteAllOrderDetailsByEventId nEventId
    end sub

    public function HasChanged()
        dim chgStatus : chgStatus = false

        for changedIndex = 0 to Count - 1
            if Item(changedIndex).HasChanged then
                chgStatus = true
                changedIndex = Count
            end if        
        next

        HasChanged = chgStatus
    end function

    public function IndexByItemNumber(ItemNumber)
        IndexByItemNumber = -1

        for itemNumIndex = 0 to Count - 1
            if Item(itemNumIndex).ItemNumber = ItemNumber then
                IndexByItemNumber = itemNumIndex
                itemNumIndex = Count
            end if 
        next
    end function

    public sub LoadByEventId(nEventId)
        dim odCon : set odCon = new OrderDetailDataConnection
        dim details : details = odCon.GetAllOrderDetailsByEventId(nEventId)

        if UBound(details) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(details, 2) 
                dim od : set od = new OrderDetail

                od.EventId = details(ORDERDETAIL_INDEX_EVENTID, i)
                od.LineNumber = details(ORDERDETAIL_INDEX_LINENUMBER, i)
                od.ItemNumber = details(ORDERDETAIL_INDEX_ITEMNUMBER, i)
                od.ItemDescription = details(ORDERDETAIL_INDEX_ITEMDESCRIPTION, i)
                od.Quantity = details(ORDERDETAIL_INDEX_QUANTITY, i)
                od.Price = details(ORDERDETAIL_INDEX_PRICE, i)
                od.CreateDateTime = details(ORDERDETAIL_INDEX_CREATEDATETIME, i)
                od.PaymentTransaction = details(ORDERDETAIL_INDEX_PAYMENTTRANSACTION, i)
                od.PaymentDateTime = details(ORDERDETAIL_INDEX_PAYMENTDATETIME, i)
                od.PaymentAmount = details(ORDERDETAIL_INDEX_PAYMENTAMOUNT, i)
                od.IsAdjustment = details(ORDERDETAIL_INDEX_ISADJUSTMENT, i)

                od.CopyOrignialValues

                mo_List.Add od
            next
        end if
    end sub

    public sub Remove(value)
        mo_List.Remove value
    end sub

    public sub SaveAll()
        for i = 0 to Count - 1
            Item(i).Save
        next
    end sub

    'properties
    public property get Count
        Count = mo_List.Count
    end property

    Public Default Property Get Item(index)
        set Item = mo_List(index)
    end property

    public property get IsAdjusted
        dim found : found = false

        for i = 0 to Count - 1
            if Item(i).IsAdjustment then
                found = true
                i = Count
            end if
        next

        IsAdjusted = found
    end property

    public property get TotalAdjustedAmount()
        dim total : total = 0.0
        dim adjustments : set adjustments = new ArrayList
        
        for i = 0 to Count - 1
            if Item(i).IsAdjustment then adjustments.Add Item(i).ItemNumber
        next

        for i = 0 to Count - 1
            dim adjustedRecord : adjustedRecord = false

            if Item(i).IsAdjustment then
                adjustedRecord = true
            else
                dim adjustmentFound : adjustmentFound = false

                for j = 0 to adjustments.Count - 1
                    if Item(i).ItemNumber = adjustments(j) then
                        adjustmentFound = true
                        j = adjustments.Count
                    end if
                next

                if adjustmentFound = false then adjustedRecord = true
            end if

            if adjustedRecord then total = total + Item(i).ExtendedPrice
        next

        TotalAdjustedAmount = total
    end property

    public property get TotalAmountOwed()
        dim total : total = 0.0

        for i = 0 to Count - 1
            dim detail : set detail = Item(i)

            if detail.IsAdjustment = false then total = total + cdbl(detail.ExtendedPrice)
        next

        TotalAmountOwed = total
    end property

    public property get TotalAmountPaid()
        dim total : total = 0.0

        for i = 0 to Count - 1
            dim detail : set detail = Item(i)

            total = total + cdbl(detail.PaymentAmount)
        next

        TotalAmountPaid = total
    end property
end class
%>