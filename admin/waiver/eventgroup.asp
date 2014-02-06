<!DOCTYPE html>
<!--#include file="../../classes/includelist.asp"-->
<%

%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title></title>
        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">

            function updateWaiver() {
                var selectedDate = new Date(parent.$("#txtSelectedDate").val());
                var originalPlayDate = new Date($("#txtOriginalPlayDate").val());
                var today = new Date();
                
                if(originalPlayDate.setHours(0, 0, 0, 0) < today.setHours(0, 0, 0, 0)) {
                    $("#txtMsg").html("Cannot edit the past.");
                } else {
                    $.ajax({
                    <%
                     If request.QueryString(QUERYSTRING_VAR_PLAYERLIST) = "" Then
                    %>
                        //Only process one player
                                url: '../../ajax/Waiver.asp?act=58&pd=<%= Request.Querystring("pd") %>&pdt=<%= Request.Querystring("pdt") %>&pt=<%= Request.Querystring("pt") %>&ev=' + $("#ddlEventId").val(),
                            <%
                             Else
                            %>
                            //Process Mulitple players
                            url: '../../ajax/Waiver.asp?act=63&pi=<%= Request.Querystring("pi") %>&pdt=<%= Request.Querystring("pdt") %>&pt=<%= Request.Querystring("pt") %>&ev=' + $("#ddlEventId").val(),
                        <%
                         End If
                        %>
                            error: function () { alert("problems encountered"); },
                            success: function () {
                                var wid = parent.$("#WaiverId").val();
                                var wdt = parent.$("#WaiverDate").val();
                    
                                if (selectedDate.setHours(0, 0, 0, 0) == originalPlayDate.setHours(0, 0, 0, 0)) {
                                    parent.$("#WaiverId").val("");
                                    parent.$("#WaiverDate").val(""); 
                                    parent.refreshAllPanels();
                                }else {
                                    parent.$("#WaiverId").val(wid);
                                    parent.$("#WaiverDate").val(wdt);
                                    parent.refreshAllPanels();                    }

                                parent.$.fancybox.close();
                            }
                        });
                }
            }
        </script>
    </head>

    <body>
            <div id="divEventGroups" class="center">
                <p>Please choose a group for<br />
                    the selected play date:</p>
                <% BuildEventList %>
                <br />
                <span id="txtMsg" style="color: red;"></span><br /><br />
                <button onclick="updateWaiver()">Update</button>
                <button onclick="parent.$.fancybox.close()">Cancel</button>
                <input type="hidden" id="txtOriginalPlayDate" name="txtOriginalPlayDate" value="<%= Request.Querystring("pdt") %>" />
            </div>
    </body>
</html>

<%
    sub BuildEventList()
        dim myEvents, selectedDateForGroups
        set myEvents = new ScheduledEventCollection
        
        If IsEmpty(request.QueryString(QUERYSTRING_VAR_PLAYDATE)) OR IsNULL(request.QueryString(QUERYSTRING_VAR_PLAYDATE)) OR Len(request.QueryString(QUERYSTRING_VAR_PLAYDATE)) = 0 Then
            selectedDateForGroups = Date
        Else
            selectedDateForGroups = request.QueryString(QUERYSTRING_VAR_PLAYDATE)
        End If

        myEvents.LoadByDate selectedDateForGroups

        response.Write "<select id=""ddlEventId"">" & vbcrlf
        response.Write "<option value="""">Walk up player</option>" & vbcrlf

        for i = 0 to myEvents.Count - 1
            Response.Write "<option value=""" & myEvents(i).EventId & """>" & iif(len(myEvents(i).PartyName) > 0, myEvents(i).PartyName, myEvents(i).ContactName) & "</option>" & vbcrlf
        next

        response.Write "</select>" & vbcrlf
    end sub    
%>