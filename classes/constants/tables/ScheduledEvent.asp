﻿<%
    const SCHEDULEDEVENT_COLUMN_ID = "EventId"
    const SCHEDULEDEVENT_COLUMN_DATE = "AvailableDate"
    const SCHEDULEDEVENT_COLUMN_STARTTIME = "StartTime"
    const SCHEDULEDEVENT_COLUMN_NUMBEROFPATRONS = "NumberOfPatrons"
    const SCHEDULEDEVENT_COLUMN_AGEOFPATRONS = "AgeOfPatrons"
    const SCHEDULEDEVENT_COLUMN_CONTACTEMAIL = "ContactEmailAddress"
    const SCHEDULEDEVENT_COLUMN_CONTACTPHONE = "ContactPhone"
    const SCHEDULEDEVENT_COLUMN_CONTACTNAME = "ContactName"
    const SCHEDULEDEVENT_COLUMN_EVENTTYPE = "EventTypeId"
    const SCHEDULEDEVENT_COLUMN_PARTYNAME = "PartyName"
    const SCHEDULEDEVENT_COLUMN_CONFIRMDATE = "ConfirmationDateTime"
    const SCHEDULEDEVENT_COLUMN_REMINDDATE = "ReminderDateTime"
    const SCHEDULEDEVENT_COLUMN_USERCOMMENTS = "UserComments"
    const SCHEDULEDEVENT_COLUMN_ADMINCOMMENTS = "AdminComments"
    const SCHEDULEDEVENT_COLUMN_CHECKINTIME = "CheckInTime"
    const SCHEDULEDEVENT_COLUMN_CHECKOUTTIME = "CheckOutTime"
    const SCHEDULEDEVENT_COLUMN_GUNSIZE = "GunSize"

    'used for invites
    'const SCHEDULEDEVENT_COLUMN_INVITETEMPLATE = "InviteTemplate"
    'const SCHEDULEDEVENT_COLUMN_PUBLICINVITELIST = "PublicInviteList"

    const SCHEDULEDEVENT_INDEX_ID = 0
    const SCHEDULEDEVENT_INDEX_DATE = 1
    const SCHEDULEDEVENT_INDEX_STARTTIME = 2
    const SCHEDULEDEVENT_INDEX_NUMBEROFPATRONS = 3
    const SCHEDULEDEVENT_INDEX_AGEOFPATRONS = 4
    const SCHEDULEDEVENT_INDEX_CONTACTEMAIL = 5
    const SCHEDULEDEVENT_INDEX_CONTACTPHONE = 6
    const SCHEDULEDEVENT_INDEX_CONTACTNAME = 7
    const SCHEDULEDEVENT_INDEX_EVENTTYPE = 8
    const SCHEDULEDEVENT_INDEX_PARTYNAME = 9
    const SCHEDULEDEVENT_INDEX_CONFIRMDATE = 10
    const SCHEDULEDEVENT_INDEX_REMINDDATE = 11
    const SCHEDULEDEVENT_INDEX_USERCOMMENTS = 12
    const SCHEDULEDEVENT_INDEX_ADMINCOMMENTS = 13
    const SCHEDULEDEVENT_INDEX_ACTIVE = 14
    const SCHEDULEDEVENT_INDEX_CHECKINTIME = 15
    const SCHEDULEDEVENT_INDEX_CHECKOUTTIME = 16
    const SCHEDULEDEVENT_INDEX_GUNSIZE = 17
    const SCHEDULEDEVENT_INDEX_WAIVERCOUNT = 18
    const SCHEDULEDEVENT_INDEX_DEPOSITAMOUNT = 19
    const SCHEDULEDEVENT_INDEX_CREATEDATETIME = 20
    const SCHEDULEDEVENT_INDEX_PAYMENTREMINDERDATETIME = 21
    const SCHEDULEDEVENT_INDEX_PAYMENTSTATUS = 22

    'used for invites
    'const SCHEDULEDEVENT_INDEX_INVITETEMPLATE = 19
    'const SCHEDULEDEVENT_INDEX_PUBLICINVITELIST = 20
%>