<%
class ScheduledEvent
    private mn_EventId, ms_Date, ms_Start, ms_NumberOfPatrons, _
            ms_AgeOfPatrons, ms_Email, ms_Phone, ms_Name, _
            mn_EventTypeId, mb_Loaded, ms_PartyName, _
            md_ConfirmationDate, md_ReminderDate, ms_UserComments, _
            ms_AdminComments, mb_Active, ms_CheckIn, ms_CheckOut, _
            ms_GunSize, mn_WaiverCount, mo_CustomFieldValues, _
            md_DepositAmount, mdt_CreateDate, mdt_PaymentReminderDate, _
            ms_PaymentStatus
    private ms_OrgDate, ms_OrgStart, ms_OrgNumberOfPatrons, _
            ms_OrgAgeOfPatrons, ms_OrgEmail, ms_OrgPhone, _
            ms_OrgName, mb_OrgLoaded, ms_OrgPartyName, _
            md_OrgConfirmationDate, md_OrgReminderDate, ms_OrgUserComments, _
            ms_OrgAdminComments, mb_OrgActive, ms_OrgCheckIn, ms_OrgCheckOut, _
            md_OrgDepositAmount, mdt_OrgCreateDate, mdt_OrgPaymentReminderDate, _
            ms_OrgPaymentStatus

    'ctor
    public sub Class_Initialize
        mb_Loaded = false
        mb_Active = false
        mn_WaiverCount = 0
        md_DepositAmount = 0
        set mo_CustomFieldValues = new RegistrationCustomValueCollection
    end sub

    ' methods
    private sub CopyOriginalValues()
        ms_OrgDate = ms_Date
        ms_OrgStart = ms_Start
        ms_OrgNumberOfPatrons = ms_NumberOfPatrons
        ms_OrgAgeOfPatrons = ms_AgeOfPatrons
        ms_OrgEmail = ms_Email
        ms_OrgPhone = ms_Phone
        ms_OrgName = ms_Name
        mb_OrgLoaded = mb_Loaded
        ms_OrgPartyName = ms_PartyName
        md_OrgConfirmationDate = md_ConfirmationDate
        md_OrgReminderDate = md_ReminderDate
        ms_OrgUserComments = ms_UserComments
        ms_OrgAdminComments = ms_AdminComments
        mb_OrgActive = mb_Active
        ms_OrgCheckIn = ms_CheckIn
        ms_OrgCheckOut = ms_CheckOut
        md_OrgDepositAmount = md_DepositAmount
        mdt_OrgCreateDate = mdt_CreateDate
        mdt_OrgPaymentReminderDate = mdt_PaymentReminderDate
        ms_OrgPaymentStatus = ms_PaymentStatus
    end sub

    public sub Delete(nEventId)
        dim myCon
        set myCon = new ScheduledEventDataConnection

        myCon.DeleteScheduledEvent(nEventId)
    end sub

    public function HasChanged()
        dim status
        status = false

        if ms_OrgDate <> ms_Date then status = true
        if ms_OrgStart <> ms_Start then status = true
        if ms_OrgNumberOfPatrons <> ms_NumberOfPatrons then status = true
        if ms_OrgAgeOfPatrons <> ms_AgeOfPatrons then status = true
        if ms_OrgEmail <> ms_Email then status = true
        if ms_OrgPhone <> ms_Phone then status = true
        if ms_OrgName <> ms_Name then status = true
        if mb_OrgLoaded <> mb_Loaded then status = true
        if ms_OrgPartyName <> ms_PartyName then status = true
        if md_OrgConfirmationDate <> md_ConfirmationDate then status = true
        if md_OrgReminderDate <> md_ReminderDate then status = true
        if ms_OrgUserComments <> ms_UserComments then status = true
        if ms_OrgAdminComments <> ms_AdminComments then status = true
        if mb_OrgActive <> mb_Active then status = true
        if ms_OrgCheckIn <> ms_CheckIn then status = true
        if ms_OrgCheckOut <> ms_CheckOut then status = true
        if cdbl(md_OrgDepositAmount) <> cdbl(md_DepositAmount) then status = true
        if mdt_OrgCreateDate <> mdt_CreateDate then status = true
        if mdt_OrgPaymentReminderDate <> mdt_PaymentReminderDate then status = true
        if ms_OrgPaymentStatus <> ms_PaymentStatus then status = true

        HasChanged = status
    end function

    public function IsConflictingSchedule()
        dim isConflict : isConflict = false
        dim concurrentEvents : set concurrentEvents = new ScheduledEventCollection
        concurrentEvents.LoadByDateAndTime SelectedDate, StartTime

        ' if there are more players than the max at the time, and there are events that are unpaid
        if concurrentEvents.TotalPatrons >= cint(Settings(SETTING_REGISTRATION_MAXPLAYERS)) and concurrentEvents.TotalUnpaidPatrons > 0 then
            for i = 0 to concurrentEvents.Count - 1
                dim myPayment : set myPayment = new ScheduledEventPayment
                myPayment.LoadLastPaymentByEventId concurrentEvents(i).EventId

                if myPayment.IsPaid() = false then
                    isConflict = true
                    exit for
                end if
            next 
        end if

        IsConflictingSchedule = isConflict
    end function

    public sub Load(nEventId)
        dim myCon
        set myCon = new ScheduledEventDataConnection

        dim results
        results = myCon.GetScheduledEvent(nEventId)

        if UBound(results) > 0 then
            EventId = results(SCHEDULEDEVENT_INDEX_ID , 0)
            SelectedDate = results(SCHEDULEDEVENT_INDEX_DATE , 0)
            StartTime = results(SCHEDULEDEVENT_INDEX_STARTTIME , 0)
            NumberOfPatrons = results(SCHEDULEDEVENT_INDEX_NUMBEROFPATRONS , 0)
            AgeOfPatrons = results(SCHEDULEDEVENT_INDEX_AGEOFPATRONS , 0)
            ContactEmailAddress = results(SCHEDULEDEVENT_INDEX_CONTACTEMAIL , 0)
            ContactPhone = results(SCHEDULEDEVENT_INDEX_CONTACTPHONE , 0)
            ContactName = results(SCHEDULEDEVENT_INDEX_CONTACTNAME , 0)
            'EventTypeId = results(SCHEDULEDEVENT_INDEX_EVENTTYPE , 0)
            PartyName = results(SCHEDULEDEVENT_INDEX_PARTYNAME , 0)
            ConfirmationDateTime = results(SCHEDULEDEVENT_INDEX_CONFIRMDATE , 0)
            ReminderDateTime = results(SCHEDULEDEVENT_INDEX_REMINDDATE , 0)
            UserComments = results(SCHEDULEDEVENT_INDEX_USERCOMMENTS , 0)
            AdminComments = results(SCHEDULEDEVENT_INDEX_ADMINCOMMENTS , 0)
            mb_Active = results(SCHEDULEDEVENT_INDEX_ACTIVE, 0)
            CheckInTime = results(SCHEDULEDEVENT_INDEX_CHECKINTIME, 0)
            CheckOutTime = results(SCHEDULEDEVENT_INDEX_CHECKOUTTIME, 0)
            'GunSize = results(SCHEDULEDEVENT_INDEX_GUNSIZE, 0)
            WaiverCount = results(SCHEDULEDEVENT_INDEX_WAIVERCOUNT, 0)
            DepositAmount = results(SCHEDULEDEVENT_INDEX_DEPOSITAMOUNT, 0)
            CreateDateTime = results(SCHEDULEDEVENT_INDEX_CREATEDATETIME, 0)
            PaymentReminderDateTime = results(SCHEDULEDEVENT_INDEX_PAYMENTREMINDERDATETIME, 0)
            PaymentStatusId = results(SCHEDULEDEVENT_INDEX_PAYMENTSTATUS, 0)

            mb_Loaded = true
        end if

        CopyOriginalValues
    end sub

    public sub LoadValues()
        CustomFieldValues.LoadByEventId EventId
    end sub

    public function reflect()
        set reflect = server.CreateObject("Scripting.Dictionary")
        reflect.Add "EventId", EncodeId(EventId)
        reflect.Add "SelectedDate", SelectedDate
        reflect.Add "StartTime", GetTimeString(StartTime)
        reflect.Add "NumberOfPatrons", NumberOfPatrons
        reflect.Add "AgeOfPatrons", AgeOfPatrons
        reflect.Add "ContactEmailAddress", ContactEmailAddress
        reflect.Add "ContactPhone", ContactPhone
        reflect.Add "ContactName", ContactName
        'reflect.Add "EventTypeId", EventTypeId
        reflect.Add "PartyName", PartyName
        reflect.Add "ConfirmationDateTime", ConfirmationDateTime
        reflect.Add "ReminderDateTime", ReminderDateTime
        reflect.Add "UserComments", UserComments
        reflect.Add "AdminComments", AdminComments
        reflect.Add "Active", Active
    end function

    public sub Save()
        dim myCon
        set myCon = new ScheduledEventDataConnection

        ' remove quotes from the party name
        dim tempPartyName
        tempPartyName = Replace(PartyName, """", "'")

        EventId = myCon.SaveScheduledEvent(EventId, SelectedDate, StartTime, NumberOfPatrons, AgeOfPatrons, ContactEmailAddress, _
                                           ContactPhone, ContactName, EventTypeId, tempPartyName, ConfirmationDateTime, _
                                           ReminderDateTime, UserComments, AdminComments, CheckInTime, CheckOutTime, _
                                           DepositAmount, CreateDateTime, PaymentReminderDateTime, PaymentStatusId)
    end sub

    public sub SaveConfirmationDate(nEventId, dConfirmDate)
        dim myCon
        set myCon = new ScheduledEventDataConnection

        myCon.SaveScheduledEventConfirmation nEventId, dConfirmDate
    end sub

    public sub SavePaymentReminderDate(nEventId, dRemindDate)
        dim myCon
        set myCon = new ScheduledEventDataConnection

        myCon.SaveScheduledEventPaymentReminder nEventId, dRemindDate
    end sub

    public sub SaveReminderDate(nEventId, dRemindDate)
        dim myCon
        set myCon = new ScheduledEventDataConnection

        myCon.SaveScheduledEventReminder nEventId, dRemindDate
    end sub

    public function ToDate()
        ToDate = CDate(SelectedDate & " " & GetMilitaryTime(StartTime))
    end function

    public function ToJSON()
        ToJSON = (new JSON).toJSON_V2(Me)
    end function

    ' properties
    public property get EventId
        EventId = mn_EventId
    end property

    public property let EventId(value)
        mn_EventId = value
    end property

    public property get SelectedDate
        SelectedDate = ms_Date
    end property

    public property let SelectedDate(value)
        ms_Date = value
    end property

    public property get StartTime
        StartTime = ms_Start
    end property

    public property get StartTimeLongFormat
        StartTimeLongFormat = GetTimeString(ms_Start)
    end property

    public property get StartTimeShortFormat
        StartTimeShortFormat = GetMilitaryTime(ms_Start)
    end property

    public property let StartTime(value)
        ms_Start = value
    end property

    public property get EstimatedEndTime
        EstimatedEndTime = FormatDateTime(DateAdd("n", cint(Settings(SETTING_REGISTRATION_PARTYDURATION)) - 1, GetTimeString(StartTime)), vbShortTime)
    end property

    public property get PreviousStartTime
        PreviousStartTime = FormatDateTime(DateAdd("n", (cint(Settings(SETTING_REGISTRATION_PARTYDURATION)) - 1) * -1, GetTimeString(StartTime)), vbShortTime)
    end property

    public property get NumberOfPatrons
        if isNull(ms_NumberOfPatrons) then
            NumberOfPatrons = ""
        else
            NumberOfPatrons = ms_NumberOfPatrons
        end if
    end property

    public property let NumberOfPatrons(value)
        ms_NumberOfPatrons = value
    end property

    public property get AgeOfPatrons
        if isNull(ms_AgeOfPatrons) then
            AgeOfPatrons = ""
        else
            AgeOfPatrons = ms_AgeOfPatrons
        end if
    end property

    public property let AgeOfPatrons(value)
        ms_AgeOfPatrons = value
    end property

    public property get ContactEmailAddress
        ContactEmailAddress = ms_Email
    end property

    public property let ContactEmailAddress(value)
        ms_Email = value
    end property

    public property get ContactPhone
        ContactPhone = ms_Phone
    end property

    public property let ContactPhone(value)
        ms_Phone = value
    end property

    public property get ContactName
        ContactName = ms_Name
    end property

    public property let ContactName(value)
        ms_Name = value
    end property

    'public property get EventTypeId
    '    EventTypeId = mn_EventTypeId
    'end property

    'public property let EventTypeId(value)
    '    mn_EventTypeId = value
    'end property

    public property get Loaded
        Loaded = mb_Loaded
    end property

    public property get PartyName
        if isNull(ms_PartyName) then
            PartyName = ""
        else
            PartyName = ms_PartyName
        end if
    end property

    public property let PartyName(value)
        ms_PartyName = value
    end property

    public property get ConfirmationDateTime
        ConfirmationDateTime = md_ConfirmationDate
    end property

    public property let ConfirmationDateTime(value)
        md_ConfirmationDate = value
    end property

    public property get ReminderDateTime
        ReminderDateTime = md_ReminderDate
    end property

    public property let ReminderDateTime(value)
        md_ReminderDate = value
    end property

    public property get UserComments
        if isNull(ms_UserComments) then
            UserComments = ""
        else
            UserComments = ms_UserComments
        end if
    end property

    public property let UserComments(value)
        ms_UserComments = value
    end property

    public property get AdminComments
        AdminComments = ms_AdminComments
    end property

    public property let AdminComments(value)
        ms_AdminComments = value
    end property

    public property get Active
        Active = mb_Active
    end property

    public property get CheckInTime
        CheckInTime = ms_CheckIn
    end property

    public property get CheckInTimeLongFormat
        CheckInTimeLongFormat = GetTimeString(ms_CheckIn)
    end property

    public property get CheckInTimeShortFormat
        CheckInTimeShortFormat = GetMilitaryTime(ms_CheckIn)
    end property

    public property let CheckInTime(value)
        ms_CheckIn = value
    end property

    public property get CheckOutTime
        CheckOutTime = ms_CheckOut
    end property

    public property get CheckOutTimeLongFormat
        CheckOutTimeLongFormat = GetTimeString(ms_CheckOut)
    end property

    public property get CheckOutTimeShortFormat
        CheckOutTimeShortFormat = GetMilitaryTime(ms_CheckOut)
    end property

    public property let CheckOutTime(value)
        ms_CheckOut = value
    end property

    'public property get GunSize
    '    if isnull(ms_GunSize) then
    '        GunSize = ""
    '    else
    '        GunSize = ms_GunSize
    '    end if
    'end property

    'public property let GunSize(value)
    '    ms_GunSize = value
    'end property

    'public property get GunSizeShort
    '    if ms_GunSize = "Not Sure" then
    '        GunSizeShort = "?"
    '    elseif instr(1, ms_GunSize, "50") > 0 then
    '        GunSizeShort =  "50"
    '    elseif instr(1, ms_GunSize, "68") > 0 then
    '        GunSizeShort =  "68"
    '    else
    '        GunSizeShort =  ""
    '    end if
    'end property

    public property get WaiverCount
        WaiverCount = mn_WaiverCount
    end property

    public property let WaiverCount(value)
        mn_WaiverCount = value
    end property

    public property get DepositAmount
        DepositAmount = md_DepositAmount
    end property

    public property let DepositAmount(value)
        md_DepositAmount = value
    end property

    public property get CustomFieldValues
        set CustomFieldValues = mo_CustomFieldValues
    end property

    public property set CustomFieldValues(value)
        set mo_CustomFieldValues = value
    end property

    public property Get CreateDateTime
        CreateDateTime = mdt_CreateDate
    end property

    public property let CreateDateTime(value)
        mdt_CreateDate = value
    end property

    public property Get PaymentReminderDateTime
        PaymentReminderDateTime = mdt_PaymentReminderDate
    end property

    public property let PaymentReminderDateTime(value)
        mdt_PaymentReminderDate = value
    end property

    public property get PaymentStatusId
        PaymentStatusId = ms_PaymentStatus
    end property

    public property let PaymentStatusId(value)
        ms_PaymentStatus = value
    end property

    public property get IsPaid
        'IsPaid = CBool(IsNull(ConfirmationDateTime) = false) 
        IsPaid = CBool((PaymentStatusId = PAYMENTSTATUS_PAID or _
                        PaymentStatusId = PAYMENTSTATUS_MARKEDPAID or _ 
                        PaymentStatusId = PAYMENTSTATUS_PAIDCONFLICTED or _ 
                        PaymentStatusId = PAYMENTSTATUS_CHANGEDPAYABLE or _ 
                        PaymentStatusId = PAYMENTSTATUS_CHANGEDRECEIVABLE)) 
    end property

    public property get IsReservingTimeSlot
        dim isPending : isPending = false

        if IsNull(ConfirmationDateTime) then
            dim lockTimeCutoff : lockTimeCutoff = DateAdd("n", 15, CreateDateTime)
            dim pendingDiff : pendingDiff = DateDiff("n", Now(), lockTimeCutoff)
    
            isPending = cbool(pendingDiff >= 0)
        end if
    
        IsReservingTimeSlot = isPending
    end property

    public property get CanCountPlayers
        CanCountPlayers = cbool(IsPaid or IsReservingTimeSlot)
    end property

    public property get IsCancelled
        IsCancelled = cbool(PaymentStatusId = PAYMENTSTATUS_CANCELLED_USER or _
                            PaymentStatusId = PAYMENTSTATUS_CANCELLED_ADMIN or _
                            PaymentStatusId = PAYMENTSTATUS_CANCELLED_TIMEOUT)
    end property
end class
%>