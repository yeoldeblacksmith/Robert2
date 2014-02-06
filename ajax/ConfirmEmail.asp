<!--#include file="../classes/IncludeList.asp"-->
<%
    const ACTION_CONFIRMEMAIL_RESEND = "0"

    select case request.QueryString(QUERYSTRING_VAR_ACTION)
        case ACTION_CONFIRMEMAIL_RESEND
            ResendConfirmation Request.QueryString(QUERYSTRING_VAR_EVENTID)
    end select

    private sub ResendConfirmation(EncodedEventId)
        dim myEvent
        set myEvent = new ScheduledEvent
        myEvent.Load DecodeId(EncodedEventId)

        dim myEmail
        set myEmail = new ConfirmationEmail
        myEmail.SendByObject myEvent

        myEvent.SaveConfirmationDate myEvent.EventId, Now()
    end sub

%>