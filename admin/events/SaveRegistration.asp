<!--#include file="../../classes/IncludeList.asp"-->
<%
    const DEBUG_MODE = false

    const ITEMNUMBER_DEPOSIT = 0
    const TEXT_ADDED = " (Added)"
    const TEXT_ADJUSTMENT = " (Adjustment)"
    const TEXT_REMOVED = " (Removed)"
    dim mb_IsPaid : mb_IsPaid = false
    dim mb_IsUpdate : mb_IsUpdate = cbool(cint(Request.Form("EventId")) > 0)
    dim mb_UsesPayments : mb_UsesPayments = UsesPayments()
    dim mo_Event : set mo_Event = GetEvent(mb_IsUpdate)
    dim mo_OrderDetails : set mo_OrderDetails = new OrderDetailCollection

    SetBaseDeposit mb_UsesPayments, mb_IsUpdate, mo_Event, mo_OrderDetails
    SaveCustomFieldValues mb_IsUpdate, mo_Event, mo_OrderDetails

    if mb_UsesPayments then
        if mb_IsUpdate then
            ' merge the existing line items with the new line items
            set mo_OrderDetails = MergeOrderDetails(mo_Event.EventId, mo_OrderDetails)

            ' have to purge all of the old records before we can save the new/updated records
            mo_OrderDetails.DeleteByEventId mo_Event.EventId
        end if

        SaveOrderDetails mb_IsPaid, mb_IsUpdate, mo_Event, mo_OrderDetails
        SaveEventPaymentStatus mb_UsesPayments, mb_IsUpdate, mb_IsPaid, mo_Event, mo_OrderDetails
        if mb_IsUpdate = false then SavePayPalPaymentStatus mb_IsPaid, mo_Event
    end if

    dim preview

    if cint(settings(SETTING_DEFAULT_EVENTID)) = cint(mo_Event.EventId) then
        preview = true
        mo_Event.Delete mo_Event.EventId
    else
        preview = false
    end if

    if Request.Form("Resend") = "on" and preview = false then
        dim myEmail : set myEmail = new ConfirmationEmail
        myEmail.SendByObject mo_Event

        mo_Event.SaveConfirmationDate mo_Event.EventId, Now
    end if
%>
<!DOCTYPE html>
<html>
    <head>
        <title></title>
        <% if DEBUG_MODE = false then 
                if preview = false then %>
          <script type="text/javascript">
            parent.reloadData();
            parent.$.fancybox.close();
          </script>
        <%      else %>
          <script type="text/javascript">
            parent.$.fancybox.close();
        </script>
        <%      end if 
           end if%>
    </head>
    <body></body>
</html>
<%
    Sub AppendNewDetails(ProcessedOrderDetails, NewOrderDetails)
        for newIndex = 0 to NewOrderDetails.Count - 1
            dim wasFound : wasFound = false
            
            for processedIndex = 0 to ProcessedOrderDetails.Count - 1
                if NewOrderDetails(newIndex).ItemNumber = ProcessedOrderDetails(processedIndex).ItemNumber then
                    wasFound = true
                    processedIndex = ProcessedOrderDetails.Count
                end if
            next

            if wasFound = false then
                NewOrderDetails(newIndex).LineNumber = ProcessedOrderDetails.Count + 1
                if NewOrderDetails(newIndex).IsPaid then
                    NewOrderDetails(newIndex).IsAdjustment = true
                    NewOrderDetails(newIndex).ItemDescription = NewOrderDetails(newIndex).ItemDescription & TEXT_ADDED
                end if

                ProcessedOrderDetails.Add NewOrderDetails(newIndex)
            end if
        next
    end sub

    function GetEvent(IsUpdate)
        dim currentEvent : set currentEvent = new ScheduledEvent

        with currentEvent
            if IsUpdate then
                .Load Request.Form("EventId")
            else
                .EventId = 0
                .CreateDateTime = Now()
                .ReminderDateTime = ""
                .DepositAmount = 0.00
                .PaymentStatusId = PAYMENTSTATUS_UNPAID
            end if
            .SelectedDate = Request.Form("Date")
            .StartTime = Request.Form("Time")
            .NumberOfPatrons = Request.Form("NumberOfGuests")
            .AgeOfPatrons = Request.Form("AgeOfGuests")
            .ContactEmailAddress = Request.Form("Email")
            .ContactPhone = Request.Form("Phone")
            .ContactName = Request.Form("Name")
            .PartyName = Request.Form("PartyName")
            .UserComments = Request.Form("UserComments")
            .AdminComments = request.Form("AdminComments")

            if .HasChanged then .Save 
        end with

        set GetEvent = currentEvent
    end function

    function GetFeeDetail(EventId, LineNumber, ItemNumber, NumberOfGuests, ItemDescription, PaymentTypeId, FeeAmount)
        dim fee : set fee = new OrderDetail
        fee.EventId = cint(EventId)
        fee.LineNumber = cint(LineNumber)
        fee.ItemNumber = ItemNumber
        fee.ItemDescription = cstr(ItemDescription)
        fee.CreateDateTime = Now()

        select case CStr(PaymentTypeId)
            case PAYMENTTYPE_BYEVENT
                fee.Quantity = 1
            case PAYMENTTYPE_BYPLAYER
                fee.Quantity = cint(NumberOfGuests)
            case PAYMENTTYPE_NONE
                fee.Quantity = 1
        end select

        fee.Price = cdbl(FeeAmount)

        set GetFeeDetail = fee
    end function

    function GetItemDescription(ItemNumber)
        dim desc

        if ItemNumber = ITEMNUMBER_DEPOSIT then
            desc = ITEM_NAME_RESERVATIONFEE
        else
            dim field : set field = new RegistrationCustomField
            field.LoadById ItemNumber

            dim myOption : set myOption = new RegistrationCustomOption
            myOption.RegistrationCustomFieldId = ItemNumber
            myOption.Sequence = Request.Form("CustomField" & ItemNumber)
            myOption.Load
        
            desc = field.Name & " - " & myOption.Text
        end if

        GetItemDescription = desc
    end function

    function GetPaymentAmountForItem(ItemNumber)
        dim feeAmount

        if ItemNumber = ITEMNUMBER_DEPOSIT then
            feeAmount = Settings(SETTING_REGISTRATION_PAYMENTAMOUNT)
        else
            if len(Request.Form("CustomField" & ItemNumber)) = 0 then
                feeAmount = 0
            else
                dim myOption : set myOption = new RegistrationCustomOption
                myOption.RegistrationCustomFieldId = ItemNumber
                myOption.Sequence = Request.Form("CustomField" & ItemNumber)
                myOption.Load

                feeAmount = myOption.value
            end if
        end if

        GetPaymentAmountForItem = cdbl(feeAmount)
    end function

    function GetPaymentTypeIdForItem(ItemNumber)
        dim paymentTypeId

        if ItemNumber = ITEMNUMBER_DEPOSIT then
            paymentTypeId = Settings(SETTING_REGISTRATION_PAYMENTTYPE)
        else
            dim field : set field = new RegistrationCustomField
            field.LoadById ItemNumber

            paymentTypeId = field.PaymentTypeId
        end if

        GetPaymentTypeIdForItem = paymentTypeId
    end function

    function MergeOrderDetails(EventId, CurrentOrderDetails)
        dim mergedDetails 

        set mergedDetails = ProcessOriginalDetails(EventId, CurrentOrderDetails)
        AppendNewDetails mergedDetails, CurrentOrderDetails

        set MergeOrderDetails = mergedDetails
    end function
    
    function ProcessOriginalDetails(EventId, NewOrderDetails)
        dim originalOrderDetails : set originalOrderDetails = new OrderDetailCollection
        originalOrderDetails.LoadByEventId EventId

        dim myMergedDetails : set myMergedDetails = new OrderDetailCollection

        for i = 0 to originalOrderDetails.Count - 1
            if originalOrderDetails(i).IsPaid then
                originalOrderDetails(i).LineNumber = myMergedDetails.Count + 1
                
                myMergedDetails.Add originalOrderDetails(i)

                dim foundIndex : foundIndex = -1

                ' look for item in new details
                for j = 0 to NewOrderDetails.Count - 1
                    If NewOrderDetails(j).ItemNumber = originalOrderDetails(i).ItemNumber then
                        foundIndex = j
                        j = NewOrderDetails.Count
                    end if
                next

                ' if it is not found, create a removal fee to zero out the extended price for the item
                if foundIndex = -1 then
                    dim removalFee : set removalFee = GetFeeDetail(EventId, _
                                                                   myMergedDetails.Count + 1, _
                                                                   originalOrderDetails(i).ItemNumber, _
                                                                   originalOrderDetails(i).Quantity, _
                                                                   originalOrderDetails(i).ItemDescription & TEXT_REMOVED, _
                                                                   GetPaymentTypeIdForItem(originalOrderDetails(i).ItemNumber), _
                                                                   GetPaymentAmountForItem(originalOrderDetails(i).ItemNumber))
                    removalFee.IsAdjustment = true
    
                    myMergedDetails.Add removalFee
                else
                    dim feeAmount : feeAmount = GetPaymentAmountForItem(originalOrderDetails(i).ItemNumber)
                    
                    if originalOrderDetails(i).Quantity <> NewOrderDetails(foundIndex).Quantity or feeAmount <> originalOrderDetails(i).Price then
                        dim adjustFee : set adjustFee = GetFeeDetail(EventId, _
                                                                     myMergedDetails.Count + 1, _
                                                                     originalOrderDetails(i).ItemNumber, _
                                                                     NewOrderDetails(foundIndex).Quantity , _
                                                                     GetItemDescription(originalOrderDetails(i).ItemNumber) & TEXT_ADJUSTMENT, _
                                                                     GetPaymentTypeIdForItem(originalOrderDetails(i).ItemNumber), _
                                                                     feeAmount)
                        adjustFee.IsAdjustment = true
    
                        myMergedDetails.Add adjustFee
                    end if
                end if
            end if
        next
    
        set ProcessOriginalDetails = myMergedDetails
    end function

    sub SaveCustomFieldValues(IsUpdate, CurrentEvent, OrderDetails)
        dim fields
        set fields = new RegistrationCustomFieldCollection
        fields.LoadAll
    
        if DEBUG_MODE then 
            response.Write "DEBUG: **************************************************************<br/>"
            response.Write "DEBUG: Saving Custom Fields<br/>"
            response.Write "DEBUG: **************************************************************<br/>"
            response.Write "DEBUG: RegistrationCustomFieldCollection.Count = " & fields.Count & "<br/>"
        end if

        for i = 0 to fields.Count - 1
            if DEBUG_MODE then 
                response.Write "DEBUG: ==============================================================<br/>"
                response.Write "DEBUG: Field.RegistrationCustomFieldId = " & fields(i).RegistrationCustomFieldId & "<br/>"
                response.Write "DEBUG: Field.CustomFieldDataType.CustomFieldDataTypeId = " & fields(i).CustomFieldDataType.CustomFieldDataTypeId & "<br/>"
                response.Write "DEBUG: Field.PaymentTypeId = " & fields(i).PaymentTypeId & "<br/>"
                response.Write "DEBUG: len(CustomField) = " & len(Request.Form("CustomField" & fields(i).RegistrationCustomFieldId)) & "<br/>"
            end if
            fields(i).SaveFormValue CurrentEvent.EventId, Request.Form("CustomField" & fields(i).RegistrationCustomFieldId)
   
            ' check for fees and total them up
            if fields(i).CustomFieldDataType.CustomFieldDataTypeId = FIELDTYPE_OPTIONS and _
               cstr(fields(i).PaymentTypeId) > PAYMENTTYPE_NONE and _
               len(Request.Form("CustomField" & fields(i).RegistrationCustomFieldId)) > 0  then
                dim myOption : set myOption = new RegistrationCustomOption
                myOption.RegistrationCustomFieldId = fields(i).RegistrationCustomFieldId
                myOption.Sequence = Request.Form("CustomField" & fields(i).RegistrationCustomFieldId)
                myOption.Load

                if DEBUG_MODE then response.Write "DEBUG: len(Option.Value) = " & len(myOption.Value) & "<br/>"

                if len(myOption.Value) > 0 then 
                    dim itemDescription : itemDescription = fields(i).Name & " - " & myOption.Text
                    dim customFee : set customFee = GetFeeDetail(CurrentEvent.EventId, _
                                                                 OrderDetails.Count + 1, _
                                                                 fields(i).RegistrationCustomFieldId, _
                                                                 CurrentEvent.NumberOfPatrons, _
                                                                 itemDescription, _
                                                                 fields(i).PaymentTypeId, _
                                                                 myOption.Value)
                    OrderDetails.Add customFee
                end if
            end if
        next
    end sub

    sub SaveEventPaymentStatus(UsesPayments, IsUpdate, IsPaid, CurrentEvent, OrderDetails)
        if IsUpdate and OrderDetails.HasChanged() then
            if CurrentEvent.PaymentStatusId = PAYMENTSTATUS_PAID or CurrentEvent.PaymentStatusId = PAYMENTSTATUS_MARKEDPAID then
                if OrderDetails.TotalAdjustedAmount < OrderDetails.TotalAmountPaid then
                    CurrentEvent.PaymentStatusId = PAYMENTSTATUS_CHANGEDPAYABLE
                    CurrentEvent.CreateDateTime = Now()
                    'CurrentEvent.Save

                    IsPaid = true
                elseif OrderDetails.TotalAdjustedAmount > OrderDetails.TotalAmountPaid then
                    CurrentEvent.PaymentStatusId = PAYMENTSTATUS_CHANGEDRECEIVABLE
                    CurrentEvent.CreateDateTime = Now()
                    'CurrentEvent.Save

                    IsPaid = false
                elseif OrderDetails.TotalAdjustedAmount = OrderDetails.TotalAmountPaid then
                    IsPaid = true
                end if
            elseif (CurrentEvent.PaymentStatusId = PAYMENTSTATUS_CHANGEDPAYABLE or CurrentEvent.PaymentStatusId = PAYMENTSTATUS_CHANGEDRECEIVABLE) and _
                   OrderDetails.IsAdjusted = false then
                dim ps : set ps = new ScheduledEventPayment
                ps.LoadLastPaymentByEventId CurrentEvent.EventId

                if ps.PayPalTransactionId = PAYPAL_TRANSID_MANUAL then
                    CurrentEvent.PaymentStatusId = PAYMENTSTATUS_MARKEDPAID
                else
                    CurrentEvent.PaymentStatusId = PAYMENTSTATUS_PAID
                end if    
                CurrentEvent.CreateDateTime = Now()
            
                IsPaid = true
            end if
        elseif UsesPayments = false or IsPaid then
            CurrentEvent.PaymentStatusId = PAYMENTSTATUS_PAID
        end if

        if CurrentEvent.HasChanged() then CurrentEvent.Save
    end sub

    sub SaveOrderDetails(IsPaid, IsUpdate, CurrentEvent, OrderDetails)
        ' if there are fees, add them to the deposit and save the adjusted amount
        if OrderDetails.Count > 0 then
            if IsUpdate = false then
                CurrentEvent.DepositAmount = OrderDetails.TotalAmountOwed
                if CurrentEvent.HasChanged() then CurrentEvent.Save
            end if

            ' if there are no fees assessed, mark them as paid
            if OrderDetails.Count > 0 and OrderDetails.TotalAmountOwed = 0 then
                IsPaid = true
                
                for j = 0 to OrderDetails.Count - 1
                    if len(OrderDetails(j).PaymentDateTime) = 0 then
                        OrderDetails(j).PaymentTransaction = "No Charge At Checkout"
                        OrderDetails(j).PaymentDateTime = Now()
                        OrderDetails(j).PaymentAmount = 0
                    end if
                next
            end if

            OrderDetails.SaveAll
        end if
    end sub

    sub SavePayPalPaymentStatus(IsPaid, CurrentEvent)
        dim payment : set payment = new ScheduledEventPayment
        
        with payment
            .EventId = CurrentEvent.EventId
            .StatusDateTime = now()
            .PaymentAmount = 0.00

            if IsPaid then
                .PaymentStatus = PAYPAL_PAYMENTSTATUS_PAID
            else
                .PaymentStatus = "Requested"
            end if
            
            .Save
        end with
    end sub

    sub SetBaseDeposit(UsesPayments, IsUpdate, CurrentEvent, OrderDetails)
        if UsesPayments and Settings(SETTING_REGISTRATION_PAYMENTTYPE) > PAYMENTTYPE_NONE then 
            dim baseFee : set baseFee = GetFeeDetail(CurrentEvent.EventId, _
                                                     OrderDetails.Count + 1, _
                                                     ITEMNUMBER_DEPOSIT, _
                                                     CurrentEvent.NumberOfPatrons, _
                                                     ITEM_NAME_RESERVATIONFEE, _
                                                     Settings(SETTING_REGISTRATION_PAYMENTTYPE), _
                                                     Settings(SETTING_REGISTRATION_PAYMENTAMOUNT))
            
            OrderDetails.Add baseFee
        end if
    end sub
%>