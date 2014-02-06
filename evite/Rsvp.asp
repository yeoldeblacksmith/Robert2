<!DOCTYPE html>
<!--#include file="common/EventIncludeList2.asp" -->
<%
    dim invite
    set invite = new Invitation
    invite.Load DecodeId(Request.QueryString(QUERYSTRING_VAR_ID))
    'invite.Load Request.QueryString(QUERYSTRING_VAR_ID)

    dim myEvent
    set myEvent = new ScheduledEvent
    myEvent.Load invite.EventId
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>RSVP For Event</title>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <!--#include file="../nghead.asp"-->
        <!--#include file="../countusers/UPDATER.ASP" -->

        <meta name="google-site-verification" content="IUE9L0Evw9Ovgc86mXY-AMjCjyFXteKGTp0NiB6eE7U" />

        <link type="text/css" rel="Stylesheet" href="content/css/eventregistration.css" />

        <script type="text/javascript" src="content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="content/js/common.js"></script>
        <script type="text/javascript">
            function updateStatus(RsvpStatus) {
                var comments = escape($("#UserComments").val());
                $.ajax({
                    url: 'common/inviteajax.asp?act=12&id=' + $("#InviteId").val() + '&np=' + $("#NumberOfGuests").val() + '&ss=' + RsvpStatus + '&uc=' + comments,
                    error: function () { alert("problems encountered"); },
                    success: function () { window.location = "rsvpsubmitted.asp"; }
                });
                //window.location = "rsvpsubmitted.asp";
            }
        </script>
    </head>
    <body>
        <!--#include file="..\gatmenuhead.asp"--></blockquote>

        <input type="hidden" id="InviteId" value="<%= invite.InvitationId %>" />

        <h1 style="text-align: center">RSVP</h1>

        <div style="margin: 0 auto; width: 300px;">
            <p>You have been invited to a party! Please use the information below to RSVP for the event.</p>

            <table border="0">
                <tr>
                    <td>Party Name:</td>
                    <td><%= myEvent.PartyName %></td>
                </tr>
                <tr>
                    <td>Contact:</td>
                    <td><%= myEvent.ContactName %></td>
                </tr>
                <tr>
                    <td>Date:</td>
                    <td><%= myEvent.SelectedDate %></td>
                </tr>
                <tr>
                    <td>Time:</td>
                    <td><%= GetTimeString(myEvent.StartTime) %></td>
                </tr>
                <tr>
                    <td>Attending:</td>
                    <td>
                        <select id="NumberOfGuests">
                            <option value="1">1</option>
                            <option value="2">2</option>
                            <option value="3">3</option>
                            <option value="4">4</option>
                            <option value="5">5</option>
                            <option value="6">6</option>
                            <option value="7">7</option>
                            <option value="8">8</option>
                            <option value="9">9</option>
                            <option value="10">10</option>
                            <option value="11">11 - 15</option>
                            <option value="15">15 - 20</option>
                            <option value="20">20 - 25</option>
                            <option value="25">25 - 30</option>
                            <option value="30">30 - 40</option>
                            <option value="40">40 - 50</option>
                            <option value="50">50 - 75</option>
                            <option value="75">75 - 100</option>
                        </select>
                    </td>
                </tr>
                <tr>
                    <td>Comments:</td>
                    <td>
                        <span name="charcounter">(1000 characters available)</span>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <textarea id="UserComments" style="width: 100%" maxlength="1000" onkeyup="handleMaxLength()"></textarea>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" style="text-align: center">
                        <button type="button" onclick="updateStatus('Going')">Accept</button>
                        &nbsp;
                        <button type="button" onclick="updateStatus('Maybe')">Maybe</button>
                        &nbsp;
                        <button type="button" onclick="updateStatus('Not Going')">Decline</button>
                    </td>
                </tr>
            </table>
        </div>

        <!--#include file="../gatsplatfooter5.asp"-->
    </body>
</html>
