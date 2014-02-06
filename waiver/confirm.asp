<!--#include file="../classes/includelist.asp"-->
<%
    dim checkingIn: checkingIn = false
    If request.form("txtCheckingIn") = "true" Then checkingIn = true End If

    dim issuesExist: issuesExit = false
    for each item in request.Form
        if Left(CStr(request.form(item)),4) = "fail" Then
            issuesExist = true
        end if
    next    
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title><%= SiteInfo.Name %> Online Waiver</title>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <meta name="viewport" content="width=device-width, maximum-scale=1.0, minimum-scale=1.0" /> 

        <link type="text/css" rel="stylesheet" href="../content/css/waiver.css" />
        <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
            function sendMail(waiverId) {
                $.ajax({
                    url: '../ajax/waiver.asp?act=50&id=' + waiverId,
                    error: function () { alert("problems encountered"); },
                    success: function () { alert("Waiver sent successfully."); }
                });
            }

            function sendAllMail(parentPlayerId) {
                $.ajax({
                    url: '../ajax/waiver.asp?act=52&pd=' + parentPlayerId,
                    error: function () { alert("problems encountered"); },
                    success: function () { alert("Waiver sent successfully."); }
                });
            }


        </script>
    </head>

    <body>
        <table class="blockCenter" style="width: 80%;">
            <tr>
                <td class="inlineCenter">
                <center> <a href="<%= SiteInfo.HomeUrl %>"><img src="<%= LOGO_URL %>" border="0" alt="Go to <%= SiteInfo.Name %> Home Page."  title="Go to <%= SiteInfo.Name %> Home Page."></a> </center>  <br>
<%
If checkingIn Then
%>
                    <img src="../content/images/biggreencheck.gif" alt="" title="Accepted" hspace="15" align="left" />
                    <h1 style="text-align: center; font-size:48px">Check-In Accepted</h1>
                    <h1 style="text-align: center">Thank you for taking the time to check-in!</h1>
                    <p class="inlineLeft"><i>You do not have to print anything.</i> But if you would like a copy of your waiver, you can click on the eye logo next to the waiver below, <img src="../content/images/eye.png"> to view or print it, or you can click on the mail icon, <img src="../content/images/email_go.png"> to have a link to your waiver emailed to you.</p>
                </td>
            </tr>
            <tr>
                <td class="inlineCenter">
                    <table style="margin: auto; width: 400px">
                    <tr><td colspan="3"><hr style="width: 400px"/></td></tr>
<%
                    WriteWaiverRow "pass", request.form("txtCheckInPID"), request.form("txtCheckInWID")
Else
    If issuesExist Then
%>
                    <img src="../content/images/Question-Mark.png" alt="" title="IssuesExist" hspace="15" align="left" />
                    <h1 style="text-align: center">Thank you for taking the time to sign your waiver online!</h1>
                    <h1 style="text-align: center; font-size:24px">However, one or more issues were detected while processing your request.</h1>
                    <p class="inlineLeft">Please see below for a description of the issues that were found.</p>
                    <p class="inlineLeft">For current waivers - <i>you do not have to print anything.</i> But if you would like a copy, you can click on the eye logo next to the waiver(s) below, <img src="../content/images/eye.png"> to view or print it, or you can click on the mail icon, <img src="../content/images/email_go.png"> to have a link to your waiver emailed to you.</p>
                    
<%
    Else
%>
                    <img src="../content/images/biggreencheck.gif" alt="" title="Accepted" hspace="15" align="left" />
                    <h1 style="text-align: center; font-size:48px">Waiver(s) Accepted</h1>
                    <h1 style="text-align: center">Thank you for taking the time to sign your waiver online!</h1>
                    <p class="inlineLeft"> We have received it - so <i>you do not have to print anything.</i> But if you would like a copy, you can click on the eye logo next to the waiver(s) below, <img src="../content/images/eye.png"> to view or print it, or you can click on the mail icon, <img src="../content/images/email_go.png"> to have a link to your waiver emailed to you.</p>
                    
<%
    End If
%>
                </td>
            </tr>
        </table>

        <table style="margin: auto; width: 400px">
        <tr><td colspan="3"><hr style="width: 400px"/></td></tr>
<%
        dim tempWaiver, waiverCount
        dim AdultPID: AdultPID = Mid(request.Form("txtMsgAdult"),9,InStr(1,request.Form("txtMsgAdult"),"WID")-9)
        dim AdultWID: AdultWID = Right(request.Form("txtMsgAdult"),Len(request.Form("txtMsgAdult"))-(InStr(1,request.Form("txtMsgAdult"),"WID")+2))
        dim Minor1PID, Minor1WID, Minor2PID, Minor2WID, Minor3PID, Minor3WID, Minor4PID, Minor4WID, Minor5PID, Minor5WID, Minor6PID, Minor6WID 

        waiverCount = 0

        if request.Form("txtSelectedPath") = 0 or request.Form("txtSelectedPath") = 2 then
            if len(request.Form("txtMsgAdult")) > 0 then
                WriteWaiverRow Left(request.Form("txtMsgAdult"),4), AdultPID, AdultWID
            end if
        end if

        if request.Form("txtSelectedPath") = 1 or request.Form("txtSelectedPath") = 2 then     
            if len(request.Form("txtMsgMinor1")) > 0 then
                Minor1PID = Mid(request.Form("txtMsgMinor1"),9,InStr(1,request.Form("txtMsgMinor1"),"WID")-9)
                Minor1WID = Right(request.Form("txtMsgMinor1"),Len(request.Form("txtMsgMinor1"))-(InStr(1,request.Form("txtMsgMinor1"),"WID")+2))
                WriteWaiverRow Left(request.Form("txtMsgMinor1"),4), Minor1PID, Minor1WID
            end if
            if len(request.Form("txtMsgMinor2")) > 0 then
                Minor2PID = Mid(request.Form("txtMsgMinor2"),9,InStr(1,request.Form("txtMsgMinor2"),"WID")-9)
                Minor2WID = Right(request.Form("txtMsgMinor2"),Len(request.Form("txtMsgMinor2"))-(InStr(1,request.Form("txtMsgMinor2"),"WID")+2))
                WriteWaiverRow Left(request.Form("txtMsgMinor2"),4), Minor2PID, Minor2WID
            end if
            if len(request.Form("txtMsgMinor3")) > 0 then
                Minor3PID = Mid(request.Form("txtMsgMinor3"),9,InStr(1,request.Form("txtMsgMinor3"),"WID")-9)
                Minor3WID = Right(request.Form("txtMsgMinor3"),Len(request.Form("txtMsgMinor3"))-(InStr(1,request.Form("txtMsgMinor3"),"WID")+2))
                WriteWaiverRow Left(request.Form("txtMsgMinor3"),4), Minor3PID, Minor3WID
            end if
            if len(request.Form("txtMsgMinor4")) > 0 then
                Minor4PID = Mid(request.Form("txtMsgMinor4"),9,InStr(1,request.Form("txtMsgMinor4"),"WID")-9)
                Minor4WID = Right(request.Form("txtMsgMinor4"),Len(request.Form("txtMsgMinor4"))-(InStr(1,request.Form("txtMsgMinor4"),"WID")+2))
                WriteWaiverRow Left(request.Form("txtMsgMinor4"),4), Minor4PID, Minor4WID
            end if
            if len(request.Form("txtMsgMinor5")) > 0 then
                Minor5PID = Mid(request.Form("txtMsgMinor5"),9,InStr(1,request.Form("txtMsgMinor5"),"WID")-9)
                Minor5WID = Right(request.Form("txtMsgMinor5"),Len(request.Form("txtMsgMinor5"))-(InStr(1,request.Form("txtMsgMinor5"),"WID")+2))
                WriteWaiverRow Left(request.Form("txtMsgMinor5"),4), Minor5PID, Minor5WID
            end if
            if len(request.Form("txtMsgMinor6")) > 0 then
                Minor6PID = Mid(request.Form("txtMsgMinor6"),9,InStr(1,request.Form("txtMsgMinor6"),"WID")-9)
                Minor6WID = Right(request.Form("txtMsgMinor6"),Len(request.Form("txtMsgMinor6"))-(InStr(1,request.Form("txtMsgMinor6"),"WID")+2))
                WriteWaiverRow Left(request.Form("txtMsgMinor6"),4), Minor6PID, Minor6WID
            end if
        end if

        If waiverCount > 1 Then
            response.Write "<tr><td colspan=3>&nbsp;</td></tr>"
            response.Write "<tr><td colspan=3><hr /></td></tr>"
            response.Write "<tr>"
            response.Write "<td><b>Send all of the waivers in one email </b></td>"
            response.Write "<td>&nbsp;</td>"
            response.Write "<td>"
            response.Write "<a href=""javascript: sendAllMail('" & AdultPID & "'); ""><img src=""../content/images/email_go.png"" alt="""" title=""Email Waiver"" border=""0""/></a>"
            response.Write "</td>"
            response.Write "</tr>"
        End If

        If issuesExist Then
            response.Write "<tr><td colspan=3>&nbsp;</td></tr>"
            response.Write "<tr><td colspan=3>&nbsp;</td></tr>"
            response.Write "<tr><td colspan=3>&nbsp;</td></tr>"
            response.Write "<tr><td colspan=3>&nbsp;</td></tr>"
            response.Write "<tr><td colspan=3>&nbsp;</td></tr>"
            response.Write "<tr>"
            response.Write "<td colspan=3>For questions or concerns, please contact us at " & SiteInfo.PhoneNumber & " during normal business hours and we will be happy to assist you.</td>"
            response.Write "</tr>"
        End If
End If 
%>
        </table>
    </body>
</html>
<%
    sub WriteWaiverRow(passfail, pid, wid)
        dim displayPlayer

        if passfail = "pass" then
            set displayPlayer = new Player            

            displayPlayer.LoadById pid

            response.Write "<tr>"
            response.Write "<td style=""text-align: left;""><b>"
            response.Write displayPlayer.FirstName & " " & displayPlayer.LastName
            response.Write "</b></td>"
            response.Write "<td>"
            response.Write "<a href=""display.asp?id=" & wid & """ target=""_blank""><img src=""../content/images/eye.png"" alt="""" title=""View Waiver"" border=""0""/></a>"
            response.Write "</td>"
            response.Write "<td>"
            response.Write "<a href=""javascript: sendMail('" & wid & "'); ""><img src=""../content/images/email_go.png"" alt="""" title=""Email Waiver"" border=""0""/></a>"
            response.Write "</td>"
            response.Write "</tr>"
            If checkingIn Then
                response.write "<tr><td colspan=3>&nbsp;</td></tr>"
            Else
                response.write "<tr><td colspan=3><span style=""color: black;"">"
                response.write "Waiver successfully saved."
                response.write "</span></td></tr>"
            End If
            response.write "<tr><td colspan=3><span>"
            response.write "You have been succesfully checked-in for your event."
            response.write "</span></td></tr>"
            response.Write "<tr><td colspan=3>&nbsp;</td></tr>"

            waiverCount = waiverCount + 1
        else
            select case wid
                case "Missing" 
                    response.Write "<tr>"
                    response.Write "<td><b>"
                    response.Write pid
                    response.Write "</b></td>"
                    response.Write "<td colspan=2>&nbsp;</td>"
                    response.Write "</tr>"
                    response.write "<tr><td colspan=3><span style=""color: red;"">"
                    response.write "Multiple players found. Waiver not saved."
                    response.write "</span></td></tr>"
                    response.write "<tr><td colspan=3><span>"
                    response.write "Please use try our <a href=""defaultW2.asp?search=true"">Search</a> to help make sure we have the right player."
                    response.write "</span></td></tr>"
                    response.Write "<tr><td colspan=3>&nbsp;</td></tr>"
                case "NoParentFound"
                    set displayPlayer = new Player            

                    displayPlayer.LoadById pid

                    response.Write "<tr>"
                    response.Write "<td style=""text-align: left;""><b>"
                    response.Write displayPlayer.FirstName & " " & displayPlayer.LastName
                    response.Write "</b></td>"
                    response.Write "<td colspan=2>&nbsp;</td>"
                    response.Write "</tr>"
                    response.write "<tr><td colspan=3><span style=""color: red;"">"
                    response.write "Minor with no parent/guardian info. Waiver not saved.<br />"
                    response.write "The was a problem saving the parent/guardian record. This is required to save a minor's waiver."
                    response.write "</span></td></tr>"
                    response.Write "<tr><td colspan=3>&nbsp;</td></tr>"
                case else
                    set displayPlayer = new Player            

                    displayPlayer.LoadById pid

                    response.Write "<tr>"
                    response.Write "<td><b>"
                    response.Write displayPlayer.FirstName & " " & displayPlayer.LastName
                    response.Write "</b></td>"
                    response.Write "<td>"
                    response.Write "<a href=""display.asp?id=" & wid & """ target=""_blank""><img src=""../content/images/eye.png"" alt="""" title=""View Waiver"" border=""0""/></a>"
                    response.Write "</td>"
                    response.Write "<td>"
                    response.Write "<a href=""javascript: sendMail('" & wid & "'); ""><img src=""../content/images/email_go.png"" alt="""" title=""Email Waiver"" border=""0""/></a>"
                    response.Write "</td>"
                    response.Write "</tr>"
                    response.write "<tr><td colspan=3><span style=""color: red;"">"
                    response.write "We found an existing current waiver already on file."
                    response.write "</span></td></tr>"
                    response.write "<tr><td colspan=3><span>"
                    response.write "You have been succesfully checked-in for your event."
                    response.write "</span></td></tr>"
                    response.Write "<tr><td colspan=3>&nbsp;</td></tr>"

                    waiverCount = waiverCount + 1
                end select
        end if
    end sub    
%>