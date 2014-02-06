<!--#include file="../classes/IncludeList.asp"-->
<%
    dim mo_fso : set mo_fso = Server.CreateObject("Scripting.FileSystemObject")
    dim ms_fileName : ms_fileName = Server.MapPath("/logs/paypal/" & GetDateString() & ".log")
    dim mo_fileStream : set mo_fileStream = mo_fso.OpenTextFile(ms_fileName, 8, true)

    WriteToNotificationLog

    dim validResponse : validResponse = RespondToPayPal()

    WriteVerifcationStatusToLog validResponse
    if validResponse then
        SavePaymentStatus
        UpdateLineItems
        HandleConfirmation
    end if

    mo_fileStream.Close

    '********************************************************************
    ' convert the date into a string that will be used for the filename
    '********************************************************************
    function GetDateString()
        dim nMonth, nDay

        nMonth = Month(now)
        nDay = Day(now)
        
        GetDateString = CStr(Year(Now)) & iif(nMonth < 10, "0" & nMonth, nMonth) & iif(nDay < 10, "0" & nDay, nDay)
    end function

    sub HandleConfirmation()
        dim myEvent : set myEvent = new ScheduledEvent
        myEvent.Load Request.Form(PAYPAL_VARIABLE_EVENTID)

        if StrComp(Request.Form(PAYPAL_VARIABLE_PAYMENTSTATUS), PAYPAL_PAYMENTSTATUS_PAID, 1) = 0 then
            dim myEmail : set myEmail = new ConfirmationEmail
            myEmail.SendByObject myEvent

            myEvent.ConfirmationDateTime = Now
            myEvent.PaymentStatusId = PAYMENTSTATUS_PAID

            if len(request.form(PAYPAL_VARIABLE_COMMENTS)) > 0 then
                dim updatedComments

                if len(myEvent.UserComments) > 0 then
                    myEvent.UserComments = myEvent.UserComments & vbCrLf & vbCrLf & "PayPal Notes: " & request.form(PAYPAL_VARIABLE_COMMENTS)
                else
                    myEvent.UserComments = "PayPal Notes: " & request.form(PAYPAL_VARIABLE_COMMENTS)
                end if
            end if

            myEvent.Save

            dim conflictingEvents : set conflictingEvents = new ScheduledEventCollection
            conflictingEvents.LoadConflictingScheduled myEvent.SelectedDate, myEvent.StartTime

            for i = 0 to conflictingEvents.Count - 1
                dim conflict : set conflict = new EventConflictEmail
                conflict.SendByObject conflictingEvents(i)

                conflictingEvents(i).PaymentStatusId = PAYMENTSTATUS_CONFLICTED
                conflictingEvents(i).Save
            next
        else
            myEvent.PaymentStatusId = PAYMENTSTATUS_FAILED
            myEvent.Save

            dim failureMail : set failureMail = new PaymentFailedEmail
            failureMail.SendByObject myEvent
        end if
    end sub

    '********************************************************************
    ' send notice back to paypal that the transaction was received
    '********************************************************************
    function RespondToPayPal()
        dim httpResponse : set httpResponse = Server.CreateObject("Msxml2.ServerXMLHTTP")
        httpResponse.Open "POST", PAYPAL_IPN_ADDRESS
        httpResponse.setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        httpResponse.Send Request.Form & "&cmd=_notify-validate"

        dim success : success = false

        if httpResponse.Status <> 200 then
            ' handle error
        elseif httpResponse.responseText = "VERIFIED" then
            success = true
        end if

        RespondToPayPal = success
    end function

    '********************************************************************
    ' write the status to our database
    '********************************************************************
    sub SavePaymentStatus()
        on error resume next
        dim payment : set payment = new ScheduledEventPayment
        
        with payment
            .EventId = Request.Form(PAYPAL_VARIABLE_EVENTID)
            '.StatusDateTime = CDate(Request.Form(PAYPAL_VARIABLE_PAYMENTDATE))
            .StatusDateTime = Now()
            .PaymentAmount = CDbl(Request.Form(PAYPAL_VARIABLE_PAYMENTAMOUNT))
            .PaymentStatus = Request.Form(PAYPAL_VARIABLE_PAYMENTSTATUS)
            .PayPalTransactionId = Request.Form(PAYPAL_VARIABLE_TRANSACTIONID)
       
            .Save

            if Err.Number <> 0 then
                dim objASPError : Set objASPError = Server.GetLastError
    
                Dim strProblem
                strProblem = "ASPCode: " & Server.HTMLEncode(objASPError.ASPCode) & vbCrLf
                strProblem = strProblem & "Number: 0x" & Hex(objASPError.Number) & vbCrLf
                strProblem = strProblem & "Source: [" & Server.HTMLEncode(objASPError.Source) & "]" & vbCrLf
                strProblem = strProblem & "Category: " & Server.HTMLEncode(objASPError.Category) & vbCrLf
                strProblem = strProblem & "File: " & Server.HTMLEncode(objASPError.File) & vbCrLf
                strProblem = strProblem & "Line: " & CStr(objASPError.Line) & vbCrLf
                strProblem = strProblem & "Column: " & CStr(objASPError.Column) & vbCrLf
                strProblem = strProblem & "Description: " & Server.HTMLEncode(objASPError.Description) & vbCrLf
                strProblem = strProblem & "Err.Description: " & Server.HTMLEncode(Err.Description) & vbCrLf
                strProblem = strProblem & "ASP Description: " & Server.HTMLEncode(objASPError.ASPDescription) & vbCrLf
                strProblem = strProblem & "Server Variables: " & vbCrLf & Server.HTMLEncode(Request.ServerVariables("ALL_HTTP")) & vbCrLf
                strProblem = strProblem & "QueryString: " & Server.HTMLEncode(Request.QueryString) & vbCrLf
                strProblem = strProblem & "URL: " & Server.HTMLEncode(Request.ServerVariables("URL")) & vbCrLf
                strProblem = strProblem & "Content Type: " & Server.HTMLEncode(Request.ServerVariables("CONTENT_TYPE")) & vbCrLf
                strProblem = strProblem & "Content Length: " & Server.HTMLEncode(Request.ServerVariables("CONTENT_LENGTH")) & vbCrLf
                strProblem = strProblem & "Local Addr: " & Server.HTMLEncode(Request.ServerVariables("LOCAL_ADDR")) & vbCrLf
                strProblem = strProblem & "Remote Addr: " & Server.HTMLEncode(Request.ServerVariables("LOCAL_ADDR")) & vbCrLf
                strProblem = strProblem & "Time: " & Now & vbCrLf

                mo_fileStream.Write strProblem

                on error goto 0
            end if
        end with
    end sub

    private sub UpdateLineItems()
        dim totalLines : totalLines = cint(Request.Form(PAYPAL_VARIABLE_LINECOUNT))

        for i = 1 to totalLines
            dim detailLine : set detailLine = new OrderDetail
            detailLine.LoadByEventIdAndLineNumber Request.Form(PAYPAL_VARIABLE_EVENTID), request.Form(PAYPAL_VARIABLE_ITEMLINENUMBER & i)

            detailLine.PaymentTransaction = Request.Form(PAYPAL_VARIABLE_TRANSACTIONID)
            detailLine.PaymentDateTime = Now()
            detailLine.PaymentAmount = Request.Form(PAYPAL_VARIABLE_PAYMENTAMOUNT & "_" & i)
            detailLine.Save
        next
    end sub

    '********************************************************************
    ' write all of the form info passed from paypal to our transaction log
    '********************************************************************
    sub WriteToNotificationLog()
        dim outputString
        dim kvpIndex

        kvpIndex = 0

        for each kvp in request.Form
            if kvpIndex > 0 then outputString = outputString & "&"

            outputString = outputString & kvp & "=" & request.Form(kvp)

            kvpIndex = kvpIndex + 1
        next
    
        mo_fileStream.WriteLine(FormatDateTime(now, 2) & " " & FormatDateTime(now, 3) & " - " & outputString)
    end sub

    '********************************************************************
    ' write paypal's verification status to the log
    '********************************************************************
    sub WriteVerifcationStatusToLog(Success)
        if Success then
            mo_fileStream.WriteLine(FormatDateTime(now, 2) & " " & FormatDateTime(now, 3) & " - Verified")
        else
            mo_fileStream.WriteLine(FormatDateTime(now, 2) & " " & FormatDateTime(now, 3) & " - Invalid")
        end if
    end sub
%>

