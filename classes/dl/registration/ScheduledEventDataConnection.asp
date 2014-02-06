<%
class ScheduledEventDataConnection
    public sub DeleteScheduledEvent(EventId)
        dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "DeleteScheduledEvent"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)

        oCon.ExecuteCommand oCmd
        oCon.CloseConnection

        set oCmd = nothing
    end sub

    public sub DeleteAllTemporaryScheduledEvents()
        dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "DeleteAllTemporaryScheduledEvents"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarchar, adParamInput, 38, MY_SITE_GUID)

        oCon.ExecuteCommand oCmd
        oCon.CloseConnection

        set oCmd = nothing
    end sub

    public function GetAllActiveScheduledEventsByDate(AvailableDate)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetAllActiveScheduledEventsByDate"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("AvailableDate", adDBDate, adParamInput, 8, AvailableDate)

        oCon.GetRecordSet oCmd, oRs

        dim results
        if oRs.EOF = false then
            results = oRs.GetRows()
        else
            results = Array()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetAllActiveScheduledEventsByDate = results
    end function

    public function GetAllActiveScheduledEventsByDateAndTime(AvailableDate, StartTime)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetAllActiveScheduledEventsByDateAndTime"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("AvailableDate", adDBDate, adParamInput, 8, AvailableDate)
        oCmd.Parameters.Append oCmd.CreateParameter("StartTime", adDBTime, adParamInput, 7, GetTimeString(StartTime))

        oCon.GetRecordSet oCmd, oRs

        dim results
        if oRs.EOF = false then
            results = oRs.GetRows()
        else
            results = Array()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetAllActiveScheduledEventsByDateAndTime = results
    end function

    public function GetAllActiveScheduledEventsByDateWithWaivers(AvailableDate)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetAllActiveScheduledEventsByDateWithWaivers"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("AvailableDate", adDBDate, adParamInput, 8, AvailableDate)

        oCon.GetRecordSet oCmd, oRs

        dim results
        if oRs.EOF = false then
            results = oRs.GetRows()
        else
            results = Array()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetAllActiveScheduledEventsByDateWithWaivers = results
    end function

    public function GetAllActiveScheduledEventsByEmail(EmailAddress)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetAllActiveScheduledEventsByEmail"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("@EmailAddress", adVarChar, adParamInput, 128, EmailAddress)

        oCon.GetRecordSet oCmd, oRs

        dim results
        if oRs.EOF = false then
            results = oRs.GetRows()
        else
            results = Array()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetAllActiveScheduledEventsByEmail = results
    end function

    public function GetAllActiveScheduledEventsRemaining()
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetAllActiveScheduledEventsRemaining"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)

        oCon.GetRecordSet oCmd, oRs

        dim results
        if oRs.EOF = false then
            results = oRs.GetRows()
        else
            results = Array()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetAllActiveScheduledEventsRemaining = results
    end function

    public function GetAllScheduledEventsByDateWithoutCheckout(AvailableDate)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetAllScheduledEventsByDateWithOutCheckOut"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("AvailableDate", adDBDate, adParamInput, 8, AvailableDate)

        oCon.GetRecordSet oCmd, oRs

        dim results
        if oRs.EOF = false then
            results = oRs.GetRows()
        else
            results = Array()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetAllScheduledEventsByDateWithoutCheckout = results
    end function

    public function GetAllScheduledEventsForPaymentReminder()
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetAllScheduledEventsForPaymentReminder"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)

        oCon.GetRecordSet oCmd, oRs

        dim results
        if oRs.EOF = false then
            results = oRs.GetRows()
        else
            results = Array()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetAllScheduledEventsForPaymentReminder = results
    end function

    public function GetAllScheduledEventsForReminder(DaysOut)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetAllScheduledEventsForReminder"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("DaysOut", adInteger, adParamInput, 4, DaysOut)

        oCon.GetRecordSet oCmd, oRs

        dim results
        if oRs.EOF = false then
            results = oRs.GetRows()
        else
            results = Array()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetAllScheduledEventsForReminder = results
    end function

    public function GetAllValidScheduledEventsByDate(AvailableDate)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetAllValidScheduledEventsByDate"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("AvailableDate", adDBDate, adParamInput, 8, AvailableDate)

        oCon.GetRecordSet oCmd, oRs

        dim results
        if oRs.EOF = false then
            results = oRs.GetRows()
        else
            results = Array()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetAllValidScheduledEventsByDate = results
    end function

    public function GetScheduledEvent(EventId)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetScheduledEvent"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)

        oCon.GetRecordSet oCmd, oRs

        dim results
        if oRs.EOF = false then
            results = oRs.GetRows()
        else
            results = Array()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetScheduledEvent = results
    end function

    public function SaveScheduledEvent(EventId, AvailableDate, StartTime, NumberOfPatrons, AgeOfPatrons, ContactEmailAddress, ContactPhone, ContactName, EventTypeId, _
                                       PartyName, ConfirmDate, RemindDate, UserComments, AdminComments, CheckInTime, CheckOutTime, DepositAmount, CreateDateTime, _
                                       PaymentReminderDateTime, PaymentStatusId)
        dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "SaveScheduledEvent"
        oCmd.CommandType = adCmdStoredProc
        
        ' required parameters
        oCmd.Parameters.Append oCmd.CreateParameter("RetVal", adInteger, adParamReturnValue)
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("@AvailableDate", adDBDate, adParamInput, 8, AvailableDate)
        oCmd.Parameters.Append oCmd.CreateParameter("@StartTime", adDBTime, adParamInput, 7, GetTimeString(StartTime))

        'optional paramters
        if IsNull(NumberOfPatrons) = false and len(NumberOfPatrons) > 0 then 
            oCmd.Parameters.Append oCmd.CreateParameter("@NumberOfPatrons", adVarChar, adParamInput, 100, NumberOfPatrons)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@NumberOfPatrons", adVarChar, adParamInput, 100, null)
        end if
        if IsNull(AgeOfPatrons) = false and len(AgeOfPatrons) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@AgeOfPatrons", adVarChar, adParamInput, 100, AgeOfPatrons)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@AgeOfPatrons", adVarChar, adParamInput, 100, null)
        end if
        if IsNull(ContactEmailAddress) = false and len(ContactEmailAddress) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@ContactEmailAddress", adVarChar, adParamInput, 128, ContactEmailAddress)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@ContactEmailAddress", adVarChar, adParamInput, 128, null)
        end if
        if IsNull(ContactPhone) = false and len(ContactPhone) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@ContactPhone", adVarChar, adParamInput, 20, ContactPhone)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@ContactPhone", adVarChar, adParamInput, 20, null)
        end if
        if IsNull(ContactName) = false and len(ContactName) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@ContactName", adVarChar, adParamInput, 100, ContactName)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@ContactName", adVarChar, adParamInput, 100, null)
        end if
        if IsNull(EventTypeId) = false and len(EventTypeId) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@EventTypeId", adInteger, adParamInput, 4, EventTypeId)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@EventTypeId", adInteger, adParamInput, 4, null)
        end if
        if IsNull(PartyName) = false and len(PartyName) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@PartyName", adVarChar, adParamInput, 100, PartyName)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@PartyName", adVarChar, adParamInput, 100, null)
        end if
        if IsNull(ConfirmDate) = false and len(ConfirmDate) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@ConfirmationDateTime", adDBDate, adParamInput, 8, ConfirmDate)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@ConfirmationDateTime", adDBDate, adParamInput, 8, null)
        end if
        if IsNull(RemindDate) = false and len(RemindDate) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@ReminderDateTime", adDBDate, adParamInput, 8, RemindDate)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@ReminderDateTime", adDBDate, adParamInput, 8, null)
        end if
        if IsNull(UserComments) = false and len(UserComments) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@UserComments", adVarChar, adParamInput, 1000, UserComments)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@UserComments", adVarChar, adParamInput, 1000, null)
        end if
        if IsNull(AdminComments) = false and len(AdminComments) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@AdminComments", adVarChar, adParamInput, 1000, AdminComments)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@AdminComments", adVarChar, adParamInput, 1000, null)
        end if
        if IsNull(CheckInTime) = false and len(CheckInTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@CheckInTime", adDBTime, adParamInput, 7, GetTimeString(CheckInTime))
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@CheckInTime", adDBTime, adParamInput, 7, null)
        end if
        if IsNull(CheckOutTime) = false and len(CheckOutTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@CheckOutTime", adDBTime, adParamInput, 7, GetTimeString(CheckOutTime))
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@CheckOutTime", adDBTime, adParamInput, 7, null)
        end if
        if isNull(DepositAmount) = false and len(DepositAmount) > 0 then
            dim parmDepositAmount : set parmDepositAmount = oCmd.CreateParameter("@DepositAmount", adDecimal, adParamInput, 9, DepositAmount)
            parmDepositAmount.NumericScale = 2
            parmDepositAmount.Precision = 9
            oCmd.Parameters.Append parmDepositAmount
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@DepositAmount", adDecimal, adParamInput, 9, null)
        end if
        if IsNull(CreateDateTime) = false and len(CreateDateTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@CreateDateTime", adDBDate, adParamInput, 8, CreateDateTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@CreateDateTime", adDBDate, adParamInput, 8, null)
        end if
        if IsNull(PaymentReminderDateTime) = false and len(PaymentReminderDateTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@PaymentReminderDateTime", adDBDate, adParamInput, 8, PaymentReminderDateTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@PaymentReminderDateTime", adDBDate, adParamInput, 8, null)
        end if
        if IsNull(PaymentStatusId) = false and len(PaymentStatusId) > 0 then 
            oCmd.Parameters.Append oCmd.CreateParameter("@PaymentStatusId", adChar, adParamInput, 2, PaymentStatusId)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@PaymentStatusId", adChar, adParamInput, 2, null)
        end if
        
        oCon.ExecuteCommand oCmd

        SaveScheduledEvent = oCmd("RetVal")

        oCon.CloseConnection


        set oCmd = nothing
    end function

    public sub SaveScheduledEventConfirmation(EventId, ConfirmDate)
        dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "SaveScheduledEventConfirmation"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)
        oCmd.Parameters.Append oCmd.CreateParameter("@ConfirmationDateTime", adDBDate, adParamInput, 8, ConfirmDate)

        oCon.ExecuteCommand oCmd
        oCon.CloseConnection

        set oCmd = nothing
    end sub

    public sub SaveScheduledEventPaymentReminder(EventId, RemindDate)
        dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "SaveScheduledEventPaymentReminder"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)
        oCmd.Parameters.Append oCmd.CreateParameter("@ReminderDateTime", adDBDate, adParamInput, 8, RemindDate)

        oCon.ExecuteCommand oCmd
        oCon.CloseConnection

        set oCmd = nothing
    end sub

    public sub SaveScheduledEventReminder(EventId, RemindDate)
        dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "SaveScheduledEventRegistration"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)
        oCmd.Parameters.Append oCmd.CreateParameter("@ReminderDateTime", adDBDate, adParamInput, 8, RemindDate)

        oCon.ExecuteCommand oCmd
        oCon.CloseConnection

        set oCmd = nothing
    end sub
end class
%>