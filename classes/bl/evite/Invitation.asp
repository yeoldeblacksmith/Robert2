<%
class Invitation
    dim mn_InvitationId, mn_EventId, ms_Name, _
        ms_EmailAddress, ms_InviteDateTime, ms_RsvpCount, _
        ms_RsvpDateTime, ms_RsvpStatus, ms_Comments, _
        mb_PublicComments

    ' ctor
    public sub Class_Initialize
    end sub

    ' methods
    public sub Delete(InvitationId)
        dim myCon
        set myCon = new InvitationDataConnection

        myCon.DeleteInvitation(InvitationId)
    end sub

    public sub Load(nInvitationId)
        dim myCon
        set myCon = new InvitationDataConnection

        dim results
        results = myCon.GetInvitation(nInvitationId)

        if UBound(results) > 0 then
            InvitationId = results(INVITATION_INDEX_ID, 0)
            'InvitationId = results(INVITATION_INDEX_ID, 0)
            EventId = results(INVITATION_INDEX_EVENTID, 0)
            Name = results(INVITATION_INDEX_NAME, 0)
            EmailAddress = results(INVITATION_INDEX_EMAILADDRESS, 0)
            InviteDateTime = results(INVITATION_INDEX_INVITEDATETIME, 0)
            RsvpCount = results(INVITATION_INDEX_RSVPCOUNT, 0)
            RsvpDateTime = results(INVITATION_INDEX_RSVPDATETIME, 0)
            RsvpStatus = results(INVITATION_INDEX_RSVPSTATUS, 0)
            Comments = results(INVITATION_INDEX_COMMENTS, 0)
            PublicComments = results(INVITATION_INDEX_PUBLICCOMMENTS, 0)
        end if
    end sub
    
    public sub Save()
        dim myCon
        set myCon = new InvitationDataConnection

        InvitationId = myCon.SaveInvitation(InvitationId, EventId, Name, EmailAddress, InviteDateTime, RsvpCount, RsvpDateTime, RsvpStatus, Comments, PublicComments)
    end sub

    ' properties
    public property Get InvitationId
        InvitationId = mn_InvitationId
    end property

    public property Let InvitationId(value)
        mn_InvitationId = value
    end property

    public property Get EventId
        EventId = mn_EventId
    end property

    public property Let EventId(value)
        mn_EventId = value
    end property

    public property Get Name
        Name = ms_Name
    end property

    public property Let Name(value)
        ms_Name = value
    end property

    public property Get EmailAddress
        EmailAddress = ms_EmailAddress
    end property

    public property Let EmailAddress(value)
        ms_EmailAddress = value
    end property

    public property Get InviteDateTime
        InviteDateTime = ms_InviteDateTime
    end property

    public property Let InviteDateTime(value)
        ms_InviteDateTime = value
    end property

    public property Get RsvpCount
        RsvpCount = ms_RsvpCount
    end property

    public property Let RsvpCount(value)
        ms_RsvpCount = value
    end property

    public property Get RsvpDateTime
        RsvpDateTime = ms_RsvpDateTime
    end property

    public property Let RsvpDateTime(value)
        ms_RsvpDateTime = value
    end property

    public property Get RsvpStatus
        if IsNull(ms_RsvpStatus) then
            RsvpStatus = ""
        else
            RsvpStatus = ms_RsvpStatus
        end if
    end property

    public property Let RsvpStatus(value)
        ms_RsvpStatus = value
    end property

    public property Get Comments
        Comments = ms_Comments
    end property

    public property Let Comments(value)
        ms_Comments = value
    end property

    public property get PublicComments
        PublicComments = mb_PublicComments
    end property

    public property let PublicComments(value)
        mb_PublicComments = value
    end property
end class
%>