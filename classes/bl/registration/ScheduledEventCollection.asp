<%
class ScheduledEventCollection
    private mo_List

    'ctor
    public sub Class_Initialize
        set mo_List = new ArrayList
    end sub

    ' methods
    public sub Add(value)
        mo_List.Add value
    end sub

    public sub DeleteTemporaryEvents()
        dim myCon : set myCon = new ScheduledEventDataConnection
        myCon.DeleteAllTemporaryScheduledEvents
    end sub

    public sub LoadByDate(sDate)
        dim myCon
        set myCon = new ScheduledEventDataConnection

        dim events
        events = myCon.GetAllActiveScheduledEventsByDate(sDate)

        if UBound(events) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(events, 2) 
                dim myEvent
                set myEvent = new ScheduledEvent

                myEvent.EventId = events(SCHEDULEDEVENT_INDEX_ID, i)
                myEvent.SelectedDate = events(SCHEDULEDEVENT_INDEX_DATE, i)
                myEvent.StartTime = events(SCHEDULEDEVENT_INDEX_STARTTIME, i)
                myEvent.NumberOfPatrons = events(SCHEDULEDEVENT_INDEX_NUMBEROFPATRONS, i)
                myEvent.AgeOfPatrons = events(SCHEDULEDEVENT_INDEX_AGEOFPATRONS, i)
                myEvent.ContactEmailAddress = events(SCHEDULEDEVENT_INDEX_CONTACTEMAIL, i)
                myEvent.ContactPhone = events(SCHEDULEDEVENT_INDEX_CONTACTPHONE, i)
                myEvent.ContactName = events(SCHEDULEDEVENT_INDEX_CONTACTNAME, i)
                'myEvent.EventTypeId = events(SCHEDULEDEVENT_INDEX_EVENTTYPE, i)
                myEvent.PartyName = events(SCHEDULEDEVENT_INDEX_PARTYNAME, i)
                myEvent.ConfirmationDateTime = events(SCHEDULEDEVENT_INDEX_CONFIRMDATE, i)
                myEvent.ReminderDateTime = events(SCHEDULEDEVENT_INDEX_REMINDDATE, i)
                myEvent.UserComments = events(SCHEDULEDEVENT_INDEX_USERCOMMENTS, i)
                myEvent.AdminComments = events(SCHEDULEDEVENT_INDEX_ADMINCOMMENTS, i)
                myEvent.CheckInTime = events(SCHEDULEDEVENT_INDEX_CHECKINTIME, i)
                myEvent.CheckOutTime = events(SCHEDULEDEVENT_INDEX_CHECKOUTTIME, i)
                'myEvent.GunSize = events(SCHEDULEDEVENT_INDEX_GUNSIZE, i)
                myEvent.WaiverCount = events(SCHEDULEDEVENT_INDEX_WAIVERCOUNT, i)
                myEvent.DepositAmount = events(SCHEDULEDEVENT_INDEX_DEPOSITAMOUNT, i)
                myEvent.CreateDateTime = events(SCHEDULEDEVENT_INDEX_CREATEDATETIME, i)
                myEvent.PaymentReminderDateTime = events(SCHEDULEDEVENT_INDEX_PAYMENTREMINDERDATETIME, i)
                myEvent.PaymentStatusId = events(SCHEDULEDEVENT_INDEX_PAYMENTSTATUS, i)

                mo_List.Add myEvent
            next
        end if
    end sub

    public sub LoadByDateAndTime(sDate, sTime)
        dim myCon
        set myCon = new ScheduledEventDataConnection

        dim events
        events = myCon.GetAllActiveScheduledEventsByDateAndTime(sDate, sTime)

        if UBound(events) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(events, 2) 
                dim myEvent
                set myEvent = new ScheduledEvent

                myEvent.EventId = events(SCHEDULEDEVENT_INDEX_ID, i)
                myEvent.SelectedDate = events(SCHEDULEDEVENT_INDEX_DATE, i)
                myEvent.StartTime = events(SCHEDULEDEVENT_INDEX_STARTTIME, i)
                myEvent.NumberOfPatrons = events(SCHEDULEDEVENT_INDEX_NUMBEROFPATRONS, i)
                myEvent.AgeOfPatrons = events(SCHEDULEDEVENT_INDEX_AGEOFPATRONS, i)
                myEvent.ContactEmailAddress = events(SCHEDULEDEVENT_INDEX_CONTACTEMAIL, i)
                myEvent.ContactPhone = events(SCHEDULEDEVENT_INDEX_CONTACTPHONE, i)
                myEvent.ContactName = events(SCHEDULEDEVENT_INDEX_CONTACTNAME, i)
                'myEvent.EventTypeId = events(SCHEDULEDEVENT_INDEX_EVENTTYPE, i)
                myEvent.PartyName = events(SCHEDULEDEVENT_INDEX_PARTYNAME, i)
                myEvent.ConfirmationDateTime = events(SCHEDULEDEVENT_INDEX_CONFIRMDATE, i)
                myEvent.ReminderDateTime = events(SCHEDULEDEVENT_INDEX_REMINDDATE, i)
                myEvent.UserComments = events(SCHEDULEDEVENT_INDEX_USERCOMMENTS, i)
                myEvent.AdminComments = events(SCHEDULEDEVENT_INDEX_ADMINCOMMENTS, i)
                myEvent.CheckInTime = events(SCHEDULEDEVENT_INDEX_CHECKINTIME, i)
                myEvent.CheckOutTime = events(SCHEDULEDEVENT_INDEX_CHECKOUTTIME, i)
                'myEvent.GunSize = events(SCHEDULEDEVENT_INDEX_GUNSIZE, i)
                myEvent.WaiverCount = events(SCHEDULEDEVENT_INDEX_WAIVERCOUNT, i)
                myEvent.DepositAmount = events(SCHEDULEDEVENT_INDEX_DEPOSITAMOUNT, i)
                myEvent.CreateDateTime = events(SCHEDULEDEVENT_INDEX_CREATEDATETIME, i)
                myEvent.PaymentReminderDateTime = events(SCHEDULEDEVENT_INDEX_PAYMENTREMINDERDATETIME, i)
                myEvent.PaymentStatusId = events(SCHEDULEDEVENT_INDEX_PAYMENTSTATUS, i)

                mo_List.Add myEvent
            next
        end if
    end sub

    public sub LoadByDateWithoutCheckout(sDate)
        dim myCon
        set myCon = new ScheduledEventDataConnection

        dim events
        events = myCon.GetAllScheduledEventsByDateWithoutCheckout(sDate)

        if UBound(events) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(events, 2) 
                dim myEvent
                set myEvent = new ScheduledEvent

                myEvent.EventId = events(SCHEDULEDEVENT_INDEX_ID, i)
                myEvent.SelectedDate = events(SCHEDULEDEVENT_INDEX_DATE, i)
                myEvent.StartTime = events(SCHEDULEDEVENT_INDEX_STARTTIME, i)
                myEvent.NumberOfPatrons = events(SCHEDULEDEVENT_INDEX_NUMBEROFPATRONS, i)
                myEvent.AgeOfPatrons = events(SCHEDULEDEVENT_INDEX_AGEOFPATRONS, i)
                myEvent.ContactEmailAddress = events(SCHEDULEDEVENT_INDEX_CONTACTEMAIL, i)
                myEvent.ContactPhone = events(SCHEDULEDEVENT_INDEX_CONTACTPHONE, i)
                myEvent.ContactName = events(SCHEDULEDEVENT_INDEX_CONTACTNAME, i)
                'myEvent.EventTypeId = events(SCHEDULEDEVENT_INDEX_EVENTTYPE, i)
                myEvent.PartyName = events(SCHEDULEDEVENT_INDEX_PARTYNAME, i)
                myEvent.ConfirmationDateTime = events(SCHEDULEDEVENT_INDEX_CONFIRMDATE, i)
                myEvent.ReminderDateTime = events(SCHEDULEDEVENT_INDEX_REMINDDATE, i)
                myEvent.UserComments = events(SCHEDULEDEVENT_INDEX_USERCOMMENTS, i)
                myEvent.AdminComments = events(SCHEDULEDEVENT_INDEX_ADMINCOMMENTS, i)
                myEvent.CheckInTime = events(SCHEDULEDEVENT_INDEX_CHECKINTIME, i)
                myEvent.CheckOutTime = events(SCHEDULEDEVENT_INDEX_CHECKOUTTIME, i)
                'myEvent.GunSize = events(SCHEDULEDEVENT_INDEX_GUNSIZE, i)
                myEvent.WaiverCount = events(SCHEDULEDEVENT_INDEX_WAIVERCOUNT, i)
                myEvent.DepositAmount = events(SCHEDULEDEVENT_INDEX_DEPOSITAMOUNT, i)
                myEvent.CreateDateTime = events(SCHEDULEDEVENT_INDEX_CREATEDATETIME, i)
                myEvent.PaymentReminderDateTime = events(SCHEDULEDEVENT_INDEX_PAYMENTREMINDERDATETIME, i)
                myEvent.PaymentStatusId = events(SCHEDULEDEVENT_INDEX_PAYMENTSTATUS, i)

                mo_List.Add myEvent
            next
        end if
    end sub

    public sub LoadByDateWithWaivers(sDate)
        dim myCon
        set myCon = new ScheduledEventDataConnection

        dim events
        events = myCon.GetAllActiveScheduledEventsByDateWithWaivers(sDate)

        if UBound(events) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(events, 2) 
                dim myEvent
                set myEvent = new ScheduledEvent

                myEvent.EventId = events(SCHEDULEDEVENT_INDEX_ID, i)
                myEvent.SelectedDate = events(SCHEDULEDEVENT_INDEX_DATE, i)
                myEvent.StartTime = events(SCHEDULEDEVENT_INDEX_STARTTIME, i)
                myEvent.NumberOfPatrons = events(SCHEDULEDEVENT_INDEX_NUMBEROFPATRONS, i)
                myEvent.AgeOfPatrons = events(SCHEDULEDEVENT_INDEX_AGEOFPATRONS, i)
                myEvent.ContactEmailAddress = events(SCHEDULEDEVENT_INDEX_CONTACTEMAIL, i)
                myEvent.ContactPhone = events(SCHEDULEDEVENT_INDEX_CONTACTPHONE, i)
                myEvent.ContactName = events(SCHEDULEDEVENT_INDEX_CONTACTNAME, i)
                'myEvent.EventTypeId = events(SCHEDULEDEVENT_INDEX_EVENTTYPE, i)
                myEvent.PartyName = events(SCHEDULEDEVENT_INDEX_PARTYNAME, i)
                myEvent.ConfirmationDateTime = events(SCHEDULEDEVENT_INDEX_CONFIRMDATE, i)
                myEvent.ReminderDateTime = events(SCHEDULEDEVENT_INDEX_REMINDDATE, i)
                myEvent.UserComments = events(SCHEDULEDEVENT_INDEX_USERCOMMENTS, i)
                myEvent.AdminComments = events(SCHEDULEDEVENT_INDEX_ADMINCOMMENTS, i)
                myEvent.CheckInTime = events(SCHEDULEDEVENT_INDEX_CHECKINTIME, i)
                myEvent.CheckOutTime = events(SCHEDULEDEVENT_INDEX_CHECKOUTTIME, i)
                'myEvent.GunSize = events(SCHEDULEDEVENT_INDEX_GUNSIZE, i)
                myEvent.WaiverCount = events(SCHEDULEDEVENT_INDEX_WAIVERCOUNT, i)
                myEvent.DepositAmount = events(SCHEDULEDEVENT_INDEX_DEPOSITAMOUNT, i)
                myEvent.CreateDateTime = events(SCHEDULEDEVENT_INDEX_CREATEDATETIME, i)
                myEvent.PaymentReminderDateTime = events(SCHEDULEDEVENT_INDEX_PAYMENTREMINDERDATETIME, i)
                myEvent.PaymentStatusId = events(SCHEDULEDEVENT_INDEX_PAYMENTSTATUS, i)

                mo_List.Add myEvent
            next
        end if
    end sub

    public sub LoadByEmail(sEmail)
        dim myCon
        set myCon = new ScheduledEventDataConnection

        dim events
        events = myCon.GetAllActiveScheduledEventsByEmail(sEmail)

        if UBound(events) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(events, 2) 
                dim myEvent
                set myEvent = new ScheduledEvent

                myEvent.EventId = events(SCHEDULEDEVENT_INDEX_ID, i)
                myEvent.SelectedDate = events(SCHEDULEDEVENT_INDEX_DATE, i)
                myEvent.StartTime = events(SCHEDULEDEVENT_INDEX_STARTTIME, i)
                myEvent.NumberOfPatrons = events(SCHEDULEDEVENT_INDEX_NUMBEROFPATRONS, i)
                myEvent.AgeOfPatrons = events(SCHEDULEDEVENT_INDEX_AGEOFPATRONS, i)
                myEvent.ContactEmailAddress = events(SCHEDULEDEVENT_INDEX_CONTACTEMAIL, i)
                myEvent.ContactPhone = events(SCHEDULEDEVENT_INDEX_CONTACTPHONE, i)
                myEvent.ContactName = events(SCHEDULEDEVENT_INDEX_CONTACTNAME, i)
                'myEvent.EventTypeId = events(SCHEDULEDEVENT_INDEX_EVENTTYPE, i)
                myEvent.PartyName = events(SCHEDULEDEVENT_INDEX_PARTYNAME, i)
                myEvent.ConfirmationDateTime = events(SCHEDULEDEVENT_INDEX_CONFIRMDATE, i)
                myEvent.ReminderDateTime = events(SCHEDULEDEVENT_INDEX_REMINDDATE, i)
                myEvent.UserComments = events(SCHEDULEDEVENT_INDEX_USERCOMMENTS, i)
                myEvent.AdminComments = events(SCHEDULEDEVENT_INDEX_ADMINCOMMENTS, i)
                myEvent.CheckInTime = events(SCHEDULEDEVENT_INDEX_CHECKINTIME, i)
                myEvent.CheckOutTime = events(SCHEDULEDEVENT_INDEX_CHECKOUTTIME, i)
                'myEvent.GunSize = events(SCHEDULEDEVENT_INDEX_GUNSIZE, i)
                myEvent.WaiverCount = events(SCHEDULEDEVENT_INDEX_WAIVERCOUNT, i)

                mo_List.Add myEvent
            next
        end if
    end sub

    public sub LoadConflictingScheduled(SelectedDate, StartTime)
        dim concurrentEvents : set concurrentEvents = new ScheduledEventCollection
        concurrentEvents.LoadByDateAndTime SelectedDate, StartTime

        ' if there are more players than the max at the time, and there are events that are unpaid
        if concurrentEvents.TotalPatrons >= cint(Settings(SETTING_REGISTRATION_MAXPLAYERS)) and concurrentEvents.TotalUnpaidPatrons > 0 then
            for i = 0 to concurrentEvents.Count - 1
                'dim myPayment : set myPayment = new ScheduledEventPayment
                'myPayment.LoadLastPaymentByEventId concurrentEvents(i).EventId

                'if myPayment.IsPaid() = false then
                if concurrentEvents(i).IsPaid = false then
                    mo_List.Add concurrentEvents(i)
                end if
            next 
        end if
    end sub

    public sub LoadForPaymentReminders()
        dim myCon
        set myCon = new ScheduledEventDataConnection

        dim results
        results = myCon.GetAllScheduledEventsForPaymentReminder()
        
        if UBound(results) > 0 then
            for i = 0 to ubound(results, 2)
                dim myEvent
                set myEvent = new ScheduledEvent

                myEvent.EventId = results(SCHEDULEDEVENT_INDEX_ID, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_DATE, i)) = false then myEvent.SelectedDate = results(SCHEDULEDEVENT_INDEX_DATE, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_STARTTIME, i)) = false then myEvent.StartTime = results(SCHEDULEDEVENT_INDEX_STARTTIME, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_NUMBEROFPATRONS, i)) = false then myEvent.NumberOfPatrons = results(SCHEDULEDEVENT_INDEX_NUMBEROFPATRONS, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_AGEOFPATRONS, i)) = false then myEvent.AgeOfPatrons = results(SCHEDULEDEVENT_INDEX_AGEOFPATRONS, i)
                myEvent.ContactEmailAddress = results(SCHEDULEDEVENT_INDEX_CONTACTEMAIL, i)
                myEvent.ContactPhone = results(SCHEDULEDEVENT_INDEX_CONTACTPHONE, i)
                myEvent.ContactName = results(SCHEDULEDEVENT_INDEX_CONTACTNAME, i)
                'if IsNull(results(SCHEDULEDEVENT_INDEX_EVENTTYPE, i)) = false then myEvent.EventTypeId = results(SCHEDULEDEVENT_INDEX_EVENTTYPE, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_PARTYNAME, i)) = false then myEvent.PartyName = results(SCHEDULEDEVENT_INDEX_PARTYNAME, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_CONFIRMDATE, i)) = false then myEvent.ConfirmationDateTime = results(SCHEDULEDEVENT_INDEX_CONFIRMDATE, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_REMINDDATE, i)) = false then myEvent.ReminderDateTime = results(SCHEDULEDEVENT_INDEX_REMINDDATE, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_USERCOMMENTS, i)) = false then myEvent.UserComments = results(SCHEDULEDEVENT_INDEX_USERCOMMENTS, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_ADMINCOMMENTS, i)) = false then myEvent.AdminComments = results(SCHEDULEDEVENT_INDEX_ADMINCOMMENTS, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_CHECKINTIME, i)) = false then myEvent.CheckInTime = results(SCHEDULEDEVENT_INDEX_CHECKINTIME, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_CHECKOUTTIME, i)) = false then myEvent.CheckOutTime = results(SCHEDULEDEVENT_INDEX_CHECKOUTTIME, i)
                'if IsNull(results(SCHEDULEDEVENT_INDEX_GUNSIZE, i)) = false then myEvent.GunSize = results(SCHEDULEDEVENT_INDEX_GUNSIZE, i)
                'if IsNull(results(SCHEDULEDEVENT_INDEX_WAIVERCOUNT, i)) = false then myEvent.WaiverCount = results(SCHEDULEDEVENT_INDEX_WAIVERCOUNT, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_DEPOSITAMOUNT, i)) = false then myEvent.DepositAmount = results(SCHEDULEDEVENT_INDEX_DEPOSITAMOUNT, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_CREATEDATETIME, i)) = false then myEvent.CreateDateTime = results(SCHEDULEDEVENT_INDEX_CREATEDATETIME, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_PAYMENTREMINDERDATETIME, i)) = false then myEvent.PaymentReminderDateTime = results(SCHEDULEDEVENT_INDEX_PAYMENTREMINDERDATETIME, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_PAYMENTSTATUS, i)) = false then myEvent.PaymentStatusId = results(SCHEDULEDEVENT_INDEX_PAYMENTSTATUS, i)

                mo_List.Add myEvent
            next
        end if
    end sub

    public sub LoadForReminders()
        dim myCon
        set myCon = new ScheduledEventDataConnection

        dim results
        results = myCon.GetAllScheduledEventsForReminder(Settings(SETTING_REGISTRATION_REMINDERDAYS))
        
        if UBound(results) > 0 then
            for i = 0 to ubound(results, 2)
                dim myEvent
                set myEvent = new ScheduledEvent

                myEvent.EventId = results(SCHEDULEDEVENT_INDEX_ID, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_DATE, i)) = false then myEvent.SelectedDate = results(SCHEDULEDEVENT_INDEX_DATE, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_STARTTIME, i)) = false then myEvent.StartTime = results(SCHEDULEDEVENT_INDEX_STARTTIME, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_NUMBEROFPATRONS, i)) = false then myEvent.NumberOfPatrons = results(SCHEDULEDEVENT_INDEX_NUMBEROFPATRONS, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_AGEOFPATRONS, i)) = false then myEvent.AgeOfPatrons = results(SCHEDULEDEVENT_INDEX_AGEOFPATRONS, i)
                myEvent.ContactEmailAddress = results(SCHEDULEDEVENT_INDEX_CONTACTEMAIL, i)
                myEvent.ContactPhone = results(SCHEDULEDEVENT_INDEX_CONTACTPHONE, i)
                myEvent.ContactName = results(SCHEDULEDEVENT_INDEX_CONTACTNAME, i)
                'if IsNull(results(SCHEDULEDEVENT_INDEX_EVENTTYPE, i)) = false then myEvent.EventTypeId = results(SCHEDULEDEVENT_INDEX_EVENTTYPE, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_PARTYNAME, i)) = false then myEvent.PartyName = results(SCHEDULEDEVENT_INDEX_PARTYNAME, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_CONFIRMDATE, i)) = false then myEvent.ConfirmationDateTime = results(SCHEDULEDEVENT_INDEX_CONFIRMDATE, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_REMINDDATE, i)) = false then myEvent.ReminderDateTime = results(SCHEDULEDEVENT_INDEX_REMINDDATE, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_USERCOMMENTS, i)) = false then myEvent.UserComments = results(SCHEDULEDEVENT_INDEX_USERCOMMENTS, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_ADMINCOMMENTS, i)) = false then myEvent.AdminComments = results(SCHEDULEDEVENT_INDEX_ADMINCOMMENTS, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_CHECKINTIME, i)) = false then myEvent.CheckInTime = results(SCHEDULEDEVENT_INDEX_CHECKINTIME, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_CHECKOUTTIME, i)) = false then myEvent.CheckOutTime = results(SCHEDULEDEVENT_INDEX_CHECKOUTTIME, i)
                'if IsNull(results(SCHEDULEDEVENT_INDEX_GUNSIZE, i)) = false then myEvent.GunSize = results(SCHEDULEDEVENT_INDEX_GUNSIZE, i)
                'if IsNull(results(SCHEDULEDEVENT_INDEX_WAIVERCOUNT, i)) = false then myEvent.WaiverCount = results(SCHEDULEDEVENT_INDEX_WAIVERCOUNT, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_DEPOSITAMOUNT, i)) = false then myEvent.DepositAmount = results(SCHEDULEDEVENT_INDEX_DEPOSITAMOUNT, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_CREATEDATETIME, i)) = false then myEvent.CreateDateTime = results(SCHEDULEDEVENT_INDEX_CREATEDATETIME, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_PAYMENTREMINDERDATETIME, i)) = false then myEvent.PaymentReminderDateTime = results(SCHEDULEDEVENT_INDEX_PAYMENTREMINDERDATETIME, i)
                if IsNull(results(SCHEDULEDEVENT_INDEX_PAYMENTSTATUS, i)) = false then myEvent.PaymentStatusId = results(SCHEDULEDEVENT_INDEX_PAYMENTSTATUS, i)

                mo_List.Add myEvent
            next
        end if
    end sub

    public sub LoadRemaining()
        dim myCon
        set myCon = new ScheduledEventDataConnection

        dim events
        events = myCon.GetAllActiveScheduledEventsRemaining()

        if UBound(events) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(events, 2) 
                dim myEvent
                set myEvent = new ScheduledEvent

                myEvent.EventId = events(SCHEDULEDEVENT_INDEX_ID, i)
                myEvent.SelectedDate = events(SCHEDULEDEVENT_INDEX_DATE, i)
                myEvent.StartTime = events(SCHEDULEDEVENT_INDEX_STARTTIME, i)
                myEvent.NumberOfPatrons = events(SCHEDULEDEVENT_INDEX_NUMBEROFPATRONS, i)
                myEvent.AgeOfPatrons = events(SCHEDULEDEVENT_INDEX_AGEOFPATRONS, i)
                myEvent.ContactEmailAddress = events(SCHEDULEDEVENT_INDEX_CONTACTEMAIL, i)
                myEvent.ContactPhone = events(SCHEDULEDEVENT_INDEX_CONTACTPHONE, i)
                myEvent.ContactName = events(SCHEDULEDEVENT_INDEX_CONTACTNAME, i)
                'myEvent.EventTypeId = events(SCHEDULEDEVENT_INDEX_EVENTTYPE, i)
                myEvent.PartyName = events(SCHEDULEDEVENT_INDEX_PARTYNAME, i)
                myEvent.ConfirmationDateTime = events(SCHEDULEDEVENT_INDEX_CONFIRMDATE, i)
                myEvent.ReminderDateTime = events(SCHEDULEDEVENT_INDEX_REMINDDATE, i)
                myEvent.UserComments = events(SCHEDULEDEVENT_INDEX_USERCOMMENTS, i)
                myEvent.AdminComments = events(SCHEDULEDEVENT_INDEX_ADMINCOMMENTS, i)
                myEvent.CheckInTime = events(SCHEDULEDEVENT_INDEX_CHECKINTIME, i)
                myEvent.CheckOutTime = events(SCHEDULEDEVENT_INDEX_CHECKOUTTIME, i)
                'myEvent.GunSize = events(SCHEDULEDEVENT_INDEX_GUNSIZE, i)
                myEvent.WaiverCount = events(SCHEDULEDEVENT_INDEX_WAIVERCOUNT, i)
                myEvent.DepositAmount = events(SCHEDULEDEVENT_INDEX_DEPOSITAMOUNT, i)
                myEvent.CreateDateTime = events(SCHEDULEDEVENT_INDEX_CREATEDATETIME, i)
                myEvent.PaymentReminderDateTime = events(SCHEDULEDEVENT_INDEX_PAYMENTREMINDERDATETIME, i)
                myEvent.PaymentStatusId = events(SCHEDULEDEVENT_INDEX_PAYMENTSTATUS, i)
    
                mo_List.Add myEvent
            next
        end if
    end sub

    public sub LoadValidByDate(sDate)
        dim myCon
        set myCon = new ScheduledEventDataConnection

        dim events
        events = myCon.GetAllValidScheduledEventsByDate(sDate)

        if UBound(events) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(events, 2) 
                dim myEvent
                set myEvent = new ScheduledEvent

                myEvent.EventId = events(SCHEDULEDEVENT_INDEX_ID, i)
                myEvent.SelectedDate = events(SCHEDULEDEVENT_INDEX_DATE, i)
                myEvent.StartTime = events(SCHEDULEDEVENT_INDEX_STARTTIME, i)
                myEvent.NumberOfPatrons = events(SCHEDULEDEVENT_INDEX_NUMBEROFPATRONS, i)
                myEvent.AgeOfPatrons = events(SCHEDULEDEVENT_INDEX_AGEOFPATRONS, i)
                myEvent.ContactEmailAddress = events(SCHEDULEDEVENT_INDEX_CONTACTEMAIL, i)
                myEvent.ContactPhone = events(SCHEDULEDEVENT_INDEX_CONTACTPHONE, i)
                myEvent.ContactName = events(SCHEDULEDEVENT_INDEX_CONTACTNAME, i)
                'myEvent.EventTypeId = events(SCHEDULEDEVENT_INDEX_EVENTTYPE, i)
                myEvent.PartyName = events(SCHEDULEDEVENT_INDEX_PARTYNAME, i)
                myEvent.ConfirmationDateTime = events(SCHEDULEDEVENT_INDEX_CONFIRMDATE, i)
                myEvent.ReminderDateTime = events(SCHEDULEDEVENT_INDEX_REMINDDATE, i)
                myEvent.UserComments = events(SCHEDULEDEVENT_INDEX_USERCOMMENTS, i)
                myEvent.AdminComments = events(SCHEDULEDEVENT_INDEX_ADMINCOMMENTS, i)
                myEvent.CheckInTime = events(SCHEDULEDEVENT_INDEX_CHECKINTIME, i)
                myEvent.CheckOutTime = events(SCHEDULEDEVENT_INDEX_CHECKOUTTIME, i)
                'myEvent.GunSize = events(SCHEDULEDEVENT_INDEX_GUNSIZE, i)
                myEvent.WaiverCount = events(SCHEDULEDEVENT_INDEX_WAIVERCOUNT, i)
                myEvent.DepositAmount = events(SCHEDULEDEVENT_INDEX_DEPOSITAMOUNT, i)
                myEvent.CreateDateTime = events(SCHEDULEDEVENT_INDEX_CREATEDATETIME, i)
                myEvent.PaymentReminderDateTime = events(SCHEDULEDEVENT_INDEX_PAYMENTREMINDERDATETIME, i)
                myEvent.PaymentStatusId = events(SCHEDULEDEVENT_INDEX_PAYMENTSTATUS, i)

                mo_List.Add myEvent
            next
        end if
    end sub

    public sub Remove(value)
        mo_List.Remove value
    end sub

    public function ToJSON()
        dim output

        output = "["
    
        for i = 0 to Count - 1
            if i > 0 then output = output & ","
            output = output & Item(i).ToJSON()
        next    

        output = output & "]"

        ToJSON = output
    end function

    'properties
    public property get Count
        Count = mo_List.Count
    end property

    Public Default Property Get Item(index)
        set Item = mo_List(index)
    end property

    public property get TotalPatrons
        dim total
        total = 0

        for i = 0 to Count - 1
            dim myEvent : set myEvent = mo_List(i)

            if len(myEvent.NumberOfPatrons) > 0 then
                if IsNumeric(myEvent.NumberOfPatrons) then
                    total = total + myEvent.NumberOfPatrons
                end if
            end if
        next

        TotalPatrons = total
    end property

    public property get TotalPaidPatrons
        dim total
        total = 0

        for i = 0 to Count - 1
            dim myEvent : set myEvent = mo_List(i)
            'dim myPayment : set myPayment = new ScheduledEventPayment

            'myPayment.LoadLastPaymentByEventId myEvent.EventId

            'if len(myEvent.NumberOfPatrons) > 0 and myPayment.IsPaid() then
            if len(myEvent.NumberOfPatrons) > 0 and _
               (myEvent.PaymentStatusId = PAYMENTSTATUS_PAID or _
                myEvent.PaymentStatusId = PAYMENTSTATUS_PAIDCONFLICTED or _ 
                myEvent.PaymentStatusId = PAYMENTSTATUS_CHANGEDPAYABLE or _ 
                myEvent.PaymentStatusId = PAYMENTSTATUS_CHANGEDRECEIVABLE) then
                if IsNumeric(myEvent.NumberOfPatrons) then
                    total = total + myEvent.NumberOfPatrons
                end if
            end if
        next

        TotalPaidPatrons = total
    end property

    public property get TotalUnpaidPatrons
        dim total
        total = 0

        for i = 0 to Count - 1
            dim myEvent : set myEvent = mo_List(i)
            'dim myPayment : set myPayment = new ScheduledEventPayment

            'myPayment.LoadLastPaymentByEventId myEvent.EventId

            'if len(myEvent.NumberOfPatrons) > 0 and myPayment.IsPaid() = false then
            'if len(myEvent.NumberOfPatrons) > 0 and myEvent.IsPaid _
            '   (myEvent.PaymentStatusId <> PAYMENTSTATUS_PAID and _
            '    myEvent.PaymentStatusId <> PAYMENTSTATUS_PAIDCONFLICTED and _ 
            '    myEvent.PaymentStatusId <> PAYMENTSTATUS_CHANGEDPAYABLE and _ 
            '    myEvent.PaymentStatusId <> PAYMENTSTATUS_CHANGEDRECEIVABLE) then
            if len(myEvent.NumberOfPatrons) > 0 and myEvent.CanCountPlayers = false and myEventIsCancelled = false then
                if IsNumeric(myEvent.NumberOfPatrons) then
                    total = total + myEvent.NumberOfPatrons
                end if
            end if
        next

        TotalUnpaidPatrons = total
    end property
end class
%>