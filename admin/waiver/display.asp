<!--#include file="../../classes/includelist.asp"-->
<%
    dim myWaiver, myWaiverCustomFields, myCustomOptions, useSvgImage, hasSingature, hasExpired, isNewAdult
    set myWaiver = new Waiver
    set myWaiverCustomFields = new WaiverCustomValueCollection
    set myCustomOptions = new WaiverCustomOptionCollection

    useSvgImage = false
    hasSingature = false
    hasExpired = false
    isNewAdult = false

    if IsNull(request.QueryString(QUERYSTRING_VAR_WAIVERID)) = false then
        myWaiver.Load request.QueryString(QUERYSTRING_VAR_WAIVERID)
        myWaiverCustomFields.LoadByWaiverId myWaiver.WaiverId
    end if

    if myWaiver.Loaded = false then
'        myWaiver.PlayDate = FormatDateTime(Now(), 2)
        myWaiver.CreateDateTime = FormatDateTime(Now(), 2)
        myWaiver.EmailList = true
'        myWaiver.EventId = Request.QueryString(QUERYSTRING_VAR_EVENTID)
        'myWaiver.DateOfBirth = "0/0/0000"
    else
        if len(myWaiver.Signature) > 0 then
            hasSingature = true
            useSvgImage = cbool(left(myWaiver.Signature, 5) = "<?xml")
        end if

        If myWaiver.ValidityStatus = WAIVER_VALIDITYSTATUS_EXPIRED Then 
            hasExpired = true
        End If

        If myWaiver.ValidityStatus = WAIVER_VALIDITYSTATUS_NEWADULT Then 
            isNewAdult = true
        End If
    end if
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title><%= SiteInfo.Name %>  Online Waiver</title>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <meta name="viewport" content="width=device-width, maximum-scale=1.0, minimum-scale=1.0" /> 

        <link rel="stylesheet" href="../../content/css/jquery.signaturepad.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/adminwaiver.css?<%=ANTI_CACHE_STRING %>" />

        <!--[if lt IE 9]><script src="../../content/js/flashcanvas.js"></script><![endif]-->
        <script src="../../content/js/jquery-1.7.2.min.js"></script>
        <script src="../../content/js/jquery.signaturepad.min.js"></script>
        <script src="../../content/js/json2.min.js"></script>

        <script type="text/javascript">
            $(document).ready(function () {
                <% if useSvgImage = false then %>
                $.ajax({
                    url: "../ajax/Waiver.asp?act=20&id=<%=myWaiver.HashId%>",
                    error: function(){ alert ("problems encountered getting player signature"); },
                    success: function(data) {
                        $(".plSig").signaturePad({ lineTop: 70, drawOnly: true, output: ".plOutput", validateFields: false, displayOnly: true }).regenerate(data);
                    }
                });
                <% end if %>

                parent.$("#PlayerId").val('<%= myWaiver.PlayerId %>');
                
            });

            function printWaiver(){
                window.open("../../waiver/display.asp?id=<%= myWaiver.HashId %>");
            }

        </script>
    </head>

    <body style="background: white">
        <% 
            Dim onClickCheckInStr: onClickCheckInStr = "onClick=""parent.checkInPlayer('display');"""
            Dim onClickCheckOutStr: onClickCheckOutStr =  "onClick=""parent.checkOutPlayer('display');"""

            If myWaiver.ValidityStatus <> WAIVER_VALIDITYSTATUS_GOOD Then 
                onClickCheckInStr = ""
                onClickCheckOutStr = ""
            End If            

        %>
        <ul class="menuBar" style="display: block">
            <li id="checkInLine" style="font-family: Verdana; font-size: 12px;" <%= onClickCheckInStr %>>
                <img src="../../content/images/door_in.png" alt="" title="Check In Players" border="0" />
                Check In
            </li>
            <li id="checkOutLine" style="font-family: Verdana; font-size: 12px;" <%= onClickCheckOutStr %>>
                <img src="../../content/images/door_out.png" alt="" title="Check Out Players" border="0" />
                Check Out
            </li>
            <li id="deleteWaiverLine" style="font-family: Verdana; font-size: 12px;" onClick="parent.deleteWaiver();">
                <img src="../../content/images/delete.png" alt="" title="Delete Waiver" border="0" />
                Delete Waiver
            </li>
            <li id="printWaiverLine" style="font-family: Verdana; font-size: 12px;" onclick="printWaiver();">
                <img src="../../content/images/printer.png" alt="" title="Print Waiver" border="0" />
                Print Waiver
            </li>
        </ul>
        <br />

        <% if hasExpired then %>
        <h1 style="text-align:center; color: red">This waiver has expired - Need one for this year</h1>
        <% elseif isNewAdult then %>
        <h1 style="text-align:center; color: red">This waiver has expired - Now an adult</h1>
        <% end if%>
        
        <table style="margin: 0 auto">
            <tr>
                <td>Participant Name:</td>
                <td>
                    <%= myWaiver.FirstName %>&nbsp;<%= myWaiver.LastName %>  &nbsp; 
                </td>
            </tr>
            <tr>
                <td>Participant Date of Birth:</td>
                <td>
                    <%= FormatDateTime(myWaiver.PlayerDOB, 2) %>
                </td>
            </tr>
            <% if len(myWaiver.ParentLastName) > 1 then %>
            <tr>
                <td>Parent/Guardian Name:</td>
                <td>
                    <%= myWaiver.ParentFirstName %> &nbsp; <%= myWaiver.ParentLastName %>
                </td>
            </tr>
                
            <% end if %>
            <tr>
                <td>Phone:</td>
                <td>
                    <%= myWaiver.PhoneNumber %>
                </td>
            </tr>
            <tr>
                <td>Email:</td>
                <td>
                    <%= myWaiver.EmailAddress %>
                </td>
            </tr>
                
            <tr>
                <td>Address:</td>
                <td>
                    <%= myWaiver.Address %> &nbsp; &nbsp; <%= myWaiver.City %> , <%= myWaiver.State %> &nbsp;<%= myWaiver.ZipCode %>
                </td>
            </tr>
            <%
                For k = 0 to myWaiverCustomFields.Count -1
                       If myWaiverCustomFields(k).LocationInWaiver <> WAIVERCUSTOMFIELDS_LOCATION_WAIVERGENERAL Then
                             with Response
                                .write "<tr>" & vbcrlf
                                .write "<td>" & myWaiverCustomFields(k).Name & "</td>" & vbcrlf

                                If myWaiverCustomFields(k).CustomFieldDataType = FIELDTYPE_OPTIONS Then
                                     myCustomOptions.LoadByFieldId myWaiverCustomFields(k).WaiverCustomFieldId 
                                     for m = 0 to myCustomOptions.Count -1
                                        if CStr(myCustomOptions(m).Sequence) = myWaiverCustomFields(k).Value then
                                            .write "<td>" & myCustomOptions(m).Text & "</td>" & vbcrlf 
                                        end if
                                     next
                                     set myCustomOptions = new WaiverCustomOptionCollection
                                Else
                                    .write "<td>" & myWaiverCustomFields(k).Value & "</td>" & vbcrlf
                                    .write "</tr>" & vbcrlf
                                End If
                             end with   
                       end if
                Next

                For k = 0 to myWaiverCustomFields.Count -1
                       If myWaiverCustomFields(k).LocationInWaiver = WAIVERCUSTOMFIELDS_LOCATION_WAIVERGENERAL Then
                             with Response
                                .write "<tr>" & vbcrlf
                                .write "<td>" & myWaiverCustomFields(k).Name & "</td>" & vbcrlf

                                If myWaiverCustomFields(k).CustomFieldDataType = FIELDTYPE_OPTIONS Then
                                     myCustomOptions.LoadByFieldId myWaiverCustomFields(k).WaiverCustomFieldId 
                                     for m = 0 to myCustomOptions.Count -1
                                        if CStr(myCustomOptions(m).Sequence) = myWaiverCustomFields(k).Value then
                                            .write "<td>" & myCustomOptions(m).Text & "</td>" & vbcrlf 
                                        end if
                                     next
                                     set myCustomOptions = new WaiverCustomOptionCollection
                                Else
                                    .write "<td>" & myWaiverCustomFields(k).Value & "</td>" & vbcrlf
                                    .write "</tr>" & vbcrlf
                                End If
                             end with   
                       end if
                Next
            %>
            <tr>
                <td>Date Signed:</td>
                <td>
                    <%= FormatDateTime(myWaiver.CreateDateTime, 2) %>
                </td>
            </tr>
            <tr>
                <td>Email list:</td>
                <td>
                    <input type="checkbox" disabled="disabled" <%= iif(cbool(myWaiver.EmailList), "checked=""checked""", "") %> />
                </td>
            </tr>
            <tr>
                <td>IP address:</td>
                <td><%= myWaiver.SubmittedFromIpAddress %></td>
            </tr>
            <% if hasSingature = false then %>
            <tr>
                <td colspan="2" style="color: red">
                    Signature does not exist
                </td>
            </tr>
            <% elseif useSvgImage then %>
            <tr>
                <td colspan="2">
                    Signature:<br />
                    <img src="../../waiver/signaturesvg.asp?id=<%= myWaiver.HashId %>" alt="Signature" />
                </td>
            </tr>
            <% else %>
            <tr>
                <td colspan="2" class="plSig">
                    Player Signature:
                    <div class="sig sigWrapper">
                        <div class="typed"></div>
                        <canvas class="pad" width="498" height="105"></canvas>
                        <input type="hidden" id="plOutput" name="plOutput" class="output"/>
                    </div>
                </td>
            </tr>
            <tr>
                <td colspan="2" class="pgSig">
                    Parent/Gaurdian Signature:
                    <div class="sig sigWrapper">
                        <div class="typed"></div>
                        <canvas class="pad" width="498" height="105"></canvas>
                        <input type="hidden" id="pgOutput" name="pgOutput" class="output"/>
                    </div>
                </td>
            </tr>
            <% end if %>
            <tr>
                <td colspan="2"><hr /></td>
            </tr>
            </table>
            
            <table style="margin: 0 auto;width: 400px;">
            <tr>
                <td>Past Play Dates:</td>
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td colspan="2">
                    <table>
<%
                        Dim myPlayDates, noPlayDatesToShow
                        noPlayDatesToShow = true
                        set myPlayDates = new PlayDateTimeCollection
                        myPlayDates.LoadAllByPlayer myWaiver.PlayerId
        
                        If myPlayDates.Count > 0 Then
                            for y = 0 to myPlayDates.Count -1
                                If (CDate(myPlayDates(y).PlayDate) < Date) AND NOT (IsNull(myPlayDates(y).CheckInDateTimeAdmin)) Then
                                    response.write "<tr><td>&nbsp;&nbsp;" & formatMyDate(myPlayDates(y).PlayDate) & " " & Left(myPlayDates(y).PlayTime,5) & "</td>"
                                    response.write "<td>" 
                                    if IsNull(myPlayDates(y).PartyName) OR IsEmpty(myPlayDates(y).PartyName) OR Len(myPlayDates(y).PartyName) <= 0 Then
                                        response.write "Walk-up Players"
                                    else
                                        response.write myPlayDates(y).PartyName
                                    end if 
                                    response.write "</td>"
                                    noPlayDatesToShow = false
                                End If
                            next
                        End If
                        If noPlayDatesToShow Then
%>
                        <tr><td style="text-align: left" colspan="2">&nbsp;&nbsp;No past play dates found</td></tr>
<%
                        End If
%>
                    </table>
                </td>
            </tr>
            <tr>
                <td>Today: </td>
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td colspan="2">
                    <table>
<%
                        set myPlayDates = nothing
                        noPlayDatesToShow = true
                        set myPlayDates = new PlayDateTimeCollection
                        myPlayDates.LoadAllByPlayer myWaiver.PlayerId

                        dim checkedInToday: checkedInToday = false
        
                        If myPlayDates.Count > 0 Then
                            for y = 0 to myPlayDates.Count -1
                                If NOT (IsNULL(myPlayDates(y).CheckInDateTimeAdmin) OR IsEmpty(myPlayDates(y).CheckInDateTimeAdmin) OR Len(myPlayDates(y).CheckInDateTimeAdmin) = 0) Then
                                    If FormatDateTime(CDate(myPlayDates(y).CheckInDateTimeAdmin),2) = FormatDateTime(Date,2) Then
                                        checkedInToday = true
                                    End If
                                End If

                                If CDate(myPlayDates(y).PlayDate) = Date Then
                                    response.write "<tr><td>"
                                    If checkedInToday Then
                                       response.write "<img src=""../../content/images/accept.png"" alt=""Checked In"" title=""Checked In"" />"
                                    End If
                                    response.write "<td/><td>&nbsp;&nbsp;" & formatMyDate(myPlayDates(y).PlayDate) & " " & Left(myPlayDates(y).PlayTime,5) & "</td>"
                                    response.write "<td>" 
                                    if IsNull(myPlayDates(y).PartyName) OR IsEmpty(myPlayDates(y).PartyName) OR Len(myPlayDates(y).PartyName) <= 0 Then
                                        response.write "Walk-up Players"
                                    else
                                        response.write myPlayDates(y).PartyName
                                    end if 
                                    response.write "</td>"
                                    noPlayDatesToShow = false
                                End If
                            next
                        End If
                        If noPlayDatesToShow Then
%>
                        <tr><td style="text-align: left" colspan="2">&nbsp;&nbsp;No play dates for today.</td></tr>
<%
                        End If
%>
                    </table>
                </td>
            </tr>



            <tr>
                <td>Upcoming Play Dates:</td>
                <td>&nbsp;</td>
            </tr>
            <tr>
                <td colspan="2">
                    <table>
<%                      set myPlayDates = nothing
                        noPlayDatesToShow = true
                        set myPlayDates = new PlayDateTimeCollection
                        myPlayDates.LoadAllByPlayer myWaiver.PlayerId
        
                        If myPlayDates.Count > 0 Then
                            for y = 0 to myPlayDates.Count -1
                                If CDate(myPlayDates(y).PlayDate) > Date Then
                                    response.write "<tr><td>"
                                    response.write "&nbsp;&nbsp;" & formatMyDate(myPlayDates(y).PlayDate) & " " & Left(myPlayDates(y).PlayTime,5) & "</td>"
                                    response.write "<td>" 
                                    if IsNull(myPlayDates(y).PartyName) OR IsEmpty(myPlayDates(y).PartyName) OR Len(myPlayDates(y).PartyName) <= 0 Then
                                        response.write "Walk-up Players"
                                    else
                                        response.write myPlayDates(y).PartyName
                                    end if 
                                    response.write "</td>"
%>
                                        <td style="text-align:right">
                                             <img src="../../content/images/calendar_delete.png" alt="" title="Delete Play Date" border="0" onClick="parent.deletePlayDate('<%= myWaiver.PlayerId %>','<%= formatMyDate(myPlayDates(y).PlayDate) %>','<%= myPlayDates(y).PlayTime %>', '<%= myWaiver.HashId %>')" />&nbsp;&nbsp;
<%
                                    If myWaiver.ValidityStatus = WAIVER_VALIDITYSTATUS_GOOD Then
%>
                                             <img src="../../content/images/calendar.png" alt="" title="Change Play Date" border="0" onClick="parent.showPlayDate('1','<%= myWaiver.PlayerId %>','<%= formatMyDate(myPlayDates(y).PlayDate) %>','<%= myPlayDates(y).PlayTime %>', '<%= myPlayDates(y).EventId  %>','<%= myPlayDates(y).CheckInDateTimeAdmin %>', '<%= myPlayDates(y).PlayDate %>', '<%= myPlayDates(y).PlayTime %>')" />&nbsp;&nbsp;
                                             <img src="../../content/images/group.png" alt="" title="Change Group" border="0" onClick="parent.showEventGroups('<%= myWaiver.PlayerId %>','<%= formatMyDate(myPlayDates(y).PlayDate) %>','<%= myPlayDates(y).PlayTime %>')" />
<%
                                    Else
                                        response.write "<td/>"
                                    End If

                                    noPlayDatesToShow = false
                                End If
                            next
                        End If
                        If noPlayDatesToShow Then
%>
                            <tr><td style="text-align: left" colspan="3">&nbsp;&nbsp;No upcoming play dates found</td></tr>
<%
                        End If
%>
<!--
                    <tr>
                        <td colspan="3"><hr /></td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <span onClick="parent.addPlayDate('<%= myWaiver.PlayerId %>','<%= myWaiver.HashId %>','2')">
                                <img src="../../content/images/calendar_add.png" alt="" title="Add New Play Date" border="0" />
                                Add New Play Date
                            </span>
                        </td>
                    </tr>
-->

                    </table>
                </td>
            </tr>
        </table>
    </body>
</html>
