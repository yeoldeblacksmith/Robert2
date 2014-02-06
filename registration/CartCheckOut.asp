<!--#include file="../classes/IncludeList.asp"-->
<%
    const ITEMNUMBER_DEPOSIT = 0
    
    dim mo_Event : set mo_Event = GetEvent()

    dim lastStatus : set lastStatus = new ScheduledEventPayment
    lastStatus.LoadLastPaymentByEventId mo_Event.EventId

    if lastStatus.PaymentStatus = PAYPAL_PAYMENTSTATUS_PAID then response.Redirect "ErrorPage.asp"

    dim mb_IsPaid : mb_IsPaid = false
    dim mo_OrderDetails : set mo_OrderDetails = new OrderDetailCollection


    SetBaseDeposit mo_Event, mo_OrderDetails
    SaveCustomFieldValues mo_Event, mo_OrderDetails

    ' have to purge all of the old records before we can save the new/updated records
    mo_OrderDetails.DeleteByEventId mo_Event.EventId

    SaveOrderDetails mb_IsPaid, mo_Event, mo_OrderDetails
    SavePayPalPaymentStatus mb_IsPaid, mo_Event


    SaveEventPaymentStatus mb_IsPaid, mo_Event, mo_OrderDetails

%>
<!DOCTYPE html>
<html>
    <head>
        <title></title>
    </head>
    <body>
        <form action="<%= PAYPAL_IPN_ADDRESS %>" method="post">
            <input type="hidden" name="<%= QUERYSTRING_VAR_ID %>" value="<%= EncodeId(mo_Event.EventId) %>" />
	        <input type="hidden" name="charset" value="utf-8"> 
	        <input type="hidden" name="cmd" value="_cart" />
            <input type="hidden" name="upload" value="1" />
            <input type="hidden" name="<%= PAYPAL_VARIABLE_EVENTID %>" value="<%= mo_Event.EventId %>" />
            <input type="hidden" name="return" value="<%= SiteInfo.VantoraUrl %>/registration/confirmation.asp" />
	        <input type="hidden" name="cancel_return" value="<%= SiteInfo.VantoraUrl %>/registration/paymentcancelled.asp?id=<%= EncodeId(mo_Event.EventId) %>" />
            <input type="hidden" name="notify_url" value="<%= SiteInfo.VantoraUrl %>/registration/paypalipnlistener.asp" />
	        <input type="hidden" name="business" value="<%= Settings(SETTING_PAYMENT_PAYPALADDRESS) %>" />
	        <input type="hidden" name="no_note" value="0" />
	        <input type="hidden" name="cbt" value="Return to <%= SiteInfo.Name %>" />
            <input type="hidden" name="currency_code" value="<%= GetCurrencyValue() %>" />
            <input type="hidden" name="rm" value="2" />

    <%
            for j = 0 to mo_OrderDetails.Count - 1
                response.Write "<input type=""hidden"" name=""" & PAYPAL_VARIABLE_ITEMDESCRIPTION & "_" & j + 1 & """ value=""" & mo_OrderDetails(j).ItemDescription & """/>" & vbCrLf
                response.Write "<input type=""hidden"" name=""" & PAYPAL_VARIABLE_ITEMLINENUMBER & "_" & j + 1 & """ value=""" & mo_OrderDetails(j).LineNumber & """/>" & vbCrLf
                response.Write "<input type=""hidden"" name=""" & PAYPAL_VARIABLE_ITEMPRICE & "_" & j + 1 & """ value=""" & FormatNumber(mo_OrderDetails(j).Price, 2) & """/>" & vbCrLf
                response.Write "<input type=""hidden"" name=""" & PAYPAL_VARIABLE_ITEMQTY & "_" & j + 1 & """ value=""" & mo_OrderDetails(j).Quantity & """/>" & vbCrLf
            next
    %>
        </form>

        <script type="text/javascript">
            document.forms[0].submit();
        </script>
    </body>
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
                NewOrderDetails(newIndex).IsAdjustment = true
                NewOrderDetails(newIndex).LineNumber = ProcessedOrderDetails.Count + 1
                NewOrderDetails(newIndex).ItemDescription = NewOrderDetails(newIndex).ItemDescription & TEXT_ADDED

                ProcessedOrderDetails.Add NewOrderDetails(newIndex)
            end if
        next
    end sub

    function GetCurrencyValue()
        select case SiteInfo.Country 
            case "CA"
                GetCurrencyValue = "CAD"
            case "US"
                GetCurrencyValue = "USD"
        end select
    end function

    function GetEvent()
        dim currentEvent : set currentEvent = new ScheduledEvent

        with currentEvent
            .Load DecodeId(Request.Form("EventId"))
            
            .SelectedDate = Request.Form("SelectedDate")
            .StartTime = Request.Form("Time")
            .NumberOfPatrons = Request.Form("NumberOfGuests")
            .AgeOfPatrons = Request.Form("AgeOfGuests")
            .ContactEmailAddress = Request.Form("Email")
            .ContactPhone = Request.Form("Phone")
            .ContactName = Request.Form("Name")
            .PartyName = Request.Form("PartyName")
            .UserComments = Request.Form("UserComments")
            .PaymentStatusId = PAYMENTSTATUS_UNPAID

            .Save 
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

    sub SaveCustomFieldValues(CurrentEvent, OrderDetails)
        dim fields
        set fields = new RegistrationCustomFieldCollection
        fields.LoadAll
    
        for i = 0 to fields.Count - 1
            fields(i).SaveFormValue CurrentEvent.EventId, Request.Form("CustomField" & fields(i).RegistrationCustomFieldId)
   
            ' check for fees and total them up
            if fields(i).CustomFieldDataType.CustomFieldDataTypeId = FIELDTYPE_OPTIONS and _
               cstr(fields(i).PaymentTypeId) > PAYMENTTYPE_NONE and _
               len(Request.Form("CustomField" & fields(i).RegistrationCustomFieldId)) > 0  then
                dim myOption : set myOption = new RegistrationCustomOption
                myOption.RegistrationCustomFieldId = fields(i).RegistrationCustomFieldId
                myOption.Sequence = Request.Form("CustomField" & fields(i).RegistrationCustomFieldId)
                myOption.Load

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

    sub SaveEventPaymentStatus(IsPaid, CurrentEvent, OrderDetails)
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
        end if

        if IsPaid then
            'dim myEmail : set myEmail = new ConfirmationEmail
            'myEmail.SendByObject mo_Event

            CurrentEvent.PaymentStatusId = PAYMENTSTATUS_PAID
            CurrentEvent.ConfirmationDateTime = Now
        end if

        CurrentEvent.Save
    end sub

    sub SaveOrderDetails(IsPaid, CurrentEvent, OrderDetails)
        ' if there are fees, add them to the deposit and save the adjusted amount
        if OrderDetails.IsAdjusted = false then
            CurrentEvent.DepositAmount = OrderDetails.TotalAmountOwed
            CurrentEvent.Save
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
    end sub

    sub SavePayPalPaymentStatus(IsPaid, CurrentEvent)
        dim statuses : set statuses = new ScheduledEventPaymentCollection
        statuses.LoadByEvent CurrentEvent.EventId

        if statuses.Count = 0 then
            dim payment : set payment = new ScheduledEventPayment
        
            with payment
                .EventId = CurrentEvent.EventId
                .StatusDateTime = now()
                .PaymentAmount = 0.00

                if IsPaid then
                    .PaymentStatus = PAYPAL_PAYMENTSTATUS_PAID
                else
                    .PaymentStatus = PAYPAL_PAYMENTSTATUS_REQUESTED
                end if
            
                .Save
            end with
        end if
    end sub

    sub SetBaseDeposit(CurrentEvent, OrderDetails)
        if Settings(SETTING_REGISTRATION_PAYMENTTYPE) > PAYMENTTYPE_NONE then 
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