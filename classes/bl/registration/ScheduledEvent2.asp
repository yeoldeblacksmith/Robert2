<%
class ScheduledEvent
    private mn_EventId, ms_Date, ms_Start, ms_NumberOfPatrons, _
            ms_AgeOfPatrons, ms_Email, ms_Phone, ms_Name, _
            mn_EventTypeId, mb_Loaded, ms_PartyName, _
            md_ConfirmationDate, md_ReminderDate, ms_UserComments, _
            ms_AdminComments, mb_Active, ms_CheckIn, ms_CheckOut, _
            ms_GunSize, ms_InviteTemplate, mb_PublicInviteList

    'ctor
    public sub Class_Initialize
        mb_Loaded = false
        mb_Active = false
    end sub

    ' methods
    public sub Delete(nEventId)
        dim myCon
        set myCon = new ScheduledEventDataConnection

        myCon.DeleteScheduledEvent(nEventId)
    end sub

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
            EventTypeId = results(SCHEDULEDEVENT_INDEX_EVENTTYPE , 0)
            PartyName = results(SCHEDULEDEVENT_INDEX_PARTYNAME , 0)
            ConfirmationDateTime = results(SCHEDULEDEVENT_INDEX_CONFIRMDATE , 0)
            ReminderDateTime = results(SCHEDULEDEVENT_INDEX_REMINDDATE , 0)
            UserComments = results(SCHEDULEDEVENT_INDEX_USERCOMMENTS , 0)
            AdminComments = results(SCHEDULEDEVENT_INDEX_ADMINCOMMENTS , 0)
            mb_Active = results(SCHEDULEDEVENT_INDEX_ACTIVE, 0)
            ms_CheckIn = results(SCHEDULEDEVENT_INDEX_CHECKINTIME, 0)
            ms_CheckOut = results(SCHEDULEDEVENT_INDEX_CHECKOUTTIME, 0)
            ms_GunSize = results(SCHEDULEDEVENT_INDEX_GUNSIZE, 0)

            mb_Loaded = true
        end if
    end sub

    public sub Save()
        dim myCon
        set myCon = new ScheduledEventDataConnection

        EventId = myCon.SaveScheduledEvent(EventId, SelectedDate, StartTime, NumberOfPatrons, AgeOfPatrons, ContactEmailAddress, _
                                           ContactPhone, ContactName, EventTypeId, PartyName, ConfirmationDateTime, _
                                           ReminderDateTime, UserComments, AdminComments, CheckInTime, CheckOutTime, _
                                           GunSize)
    end sub

    public sub SaveConfirmationDate(nEventId, dConfirmDate)
        dim myCon
        set myCon = new ScheduledEventDataConnection

        myCon.SaveScheduledEventConfirmation nEventId, dConfirmDate
    end sub

    public sub SaveReminderDate(nEventId, dRemindDate)
        dim myCon
        set myCon = new ScheduledEventDataConnection

        myCon.SaveScheduledEventReminder nEventId, dRemindDate
    end sub

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
        EstimatedEndTime = FormatDateTime(DateAdd("n", 150, GetTimeString(StartTime)), vbShortTime)
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

    public property get EventTypeId
        EventTypeId = mn_EventTypeId
    end property

    public property let EventTypeId(value)
        mn_EventTypeId = value
    end property

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

    public property get GunSize
        GunSize = ms_GunSize
    end property

    public property let GunSize(value)
        ms_GunSize = value
    end property

    public property get GunSizeShort
        if ms_GunSize = "Not Sure" then
            GunSizeShort = "?"
        elseif instr(1, ms_GunSize, "50") > 0 then
            GunSizeShort =  "50"
        elseif instr(1, ms_GunSize, "68") > 0 then
            GunSizeShort =  "68"
        else
            GunSizeShort =  ""
        end if
    end property

    public property get InviteTemplate
        InviteTemplate = ms_InviteTemplate
    end property

    public property let InviteTemplate(value)
        ms_InviteTemplate = value
    end property

    public property get PublicInviteList
        PubliceInviteList = mb_PubliceInviteList
    end property

    public property let PublicInviteList(value)
        mb_PubliceInviteList = value
    end property
end class
%>