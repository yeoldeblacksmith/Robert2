<%
class InvitationCollection
    private mo_List

    'ctor
    public sub Class_Initialize
        set mo_List = new ArrayList
    end sub

    ' methods
    public sub Add(value)
        mo_List.Add value
    end sub

    public sub LoadByEvent(EventId)
        dim myCon
        set myCon = new InvitationDataConnection

        dim results
        results = myCon.GetAllInvitationsByEventId(EventId)

        if UBound(results) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(results, 2) 
                dim myInvite
                set myInvite = new Invitation

                myInvite.InvitationId = results(INVITATION_INDEX_ID, i)
                myInvite.EventId = results(INVITATION_INDEX_EVENTID, i)
                myInvite.Name = results(INVITATION_INDEX_NAME, i)
                myInvite.EmailAddress = results(INVITATION_INDEX_EMAILADDRESS, i)
                myInvite.InviteDateTime = results(INVITATION_INDEX_INVITEDATETIME, i)
                myInvite.RsvpCount = results(INVITATION_INDEX_RSVPCOUNT, i)
                myInvite.RsvpDateTime = results(INVITATION_INDEX_RSVPDATETIME, i)
                myInvite.RsvpStatus = results(INVITATION_INDEX_RSVPSTATUS, i)
                myInvite.Comments = results(INVITATION_INDEX_COMMENTS, i)
                myInvite.PublicComments = results(INVITATION_INDEX_PUBLICCOMMENTS, i)

                mo_List.Add(myInvite)
            next
        end if
    end sub

    public sub LoadUnsentByEvent(EventId)
        dim myCon
        set myCon = new InvitationDataConnection

        dim results
        results = myCon.GetAllInvitationsUnsentByEventId(EventId)

        if UBound(results) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(results, 2) 
                dim myInvite
                set myInvite = new Invitation

                myInvite.InvitationId = results(INVITATION_INDEX_ID, i)
                myInvite.EventId = results(INVITATION_INDEX_EVENTID, i)
                myInvite.Name = results(INVITATION_INDEX_NAME, i)
                myInvite.EmailAddress = results(INVITATION_INDEX_EMAILADDRESS, i)
                myInvite.InviteDateTime = results(INVITATION_INDEX_INVITEDATETIME, i)
                myInvite.RsvpCount = results(INVITATION_INDEX_RSVPCOUNT, i)
                myInvite.RsvpDateTime = results(INVITATION_INDEX_RSVPDATETIME, i)
                myInvite.RsvpStatus = results(INVITATION_INDEX_RSVPSTATUS, i)
                myInvite.Comments = results(INVITATION_INDEX_COMMENTS, i)
                myInvite.PublicComments = results(INVITATION_INDEX_PUBLICCOMMENTS, i)

                mo_List.Add(myInvite)
            next
        end if
    end sub

    public sub Remove(value)
        mo_List.Remove value
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