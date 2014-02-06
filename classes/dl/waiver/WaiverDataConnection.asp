<%
class WaiverDataConnection

    public sub DeleteWaiver(sHashId)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_DeleteWaiver"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@HashId", adChar, adParamInput, 32, sHashId)

        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub

    '****************************************************************************
    ' return all of the waivers for the specific event
    '****************************************************************************
    public function GetWaiversByEventId(EventId, IncludeInvalid, IncludeCheckedIn)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetWaiverByEventId"
        oCmd.CommandType = adCmdStoredProc

        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)

        if IsNull(EventId) = false and len(EventId) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, null)
        end if

        oCmd.Parameters.Append oCmd.CreateParameter("@includeInvalid", adBoolean, adParamInput, , IncludeInvalid)
        oCmd.Parameters.Append oCmd.CreateParameter("@IncludeCheckedIn", adBoolean, adParamInput, , IncludeCheckedIn)

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

        GetWaiversByEventId = results
    end function

    '****************************************************************************
    ' return all of the waivers that do not belong to an event
    '****************************************************************************
    public function GetWaiversDateAndTimeWithoutEvent(PlayDate, PlayTime, IncludeInvalid, IncludeCheckedIn)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetWaiversByDateAndTimeWithoutEvent"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayDate", adDate, adParamInput, 7, PlayDate)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayTime", adDate, adParamInput, 7, PlayTime)
        oCmd.Parameters.Append oCmd.CreateParameter("@includeInvalid", adBoolean, adParamInput, , IncludeInvalid)
        oCmd.Parameters.Append oCmd.CreateParameter("@IncludeCheckedIn", adBoolean, adParamInput, , IncludeCheckedIn)
        
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

        GetWaiversDateAndTimeWithoutEvent = results
    end function

    '****************************************************************************
    ' return all of the waivers that do not belong to an event
    '****************************************************************************
    public function GetWaiversForDateWithoutEvent(PlayDate, IncludeInvalid, IncludeCheckedIn)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetWaiversForDateWithoutEvent"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayDate", adDate, adParamInput, 7, PlayDate)
        oCmd.Parameters.Append oCmd.CreateParameter("@includeInvalid", adBoolean, adParamInput, , IncludeInvalid)
        oCmd.Parameters.Append oCmd.CreateParameter("@IncludeCheckedIn", adBoolean, adParamInput, , IncludeCheckedIn)
        
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

        GetWaiversForDateWithoutEvent = results
    end function

    '****************************************************************************
    ' get waivers by playerid 
    '****************************************************************************
    public function GetWaiversByPlayerId(nPlayerId)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetWaiver2sByPlayerId"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@PassedInSiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("@PassedInPlayerId", adInteger, adParamInput, 4, nPlayerId)
        
        
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

        GetWaiversByPlayerId = results

    end function

    '****************************************************************************
    ' get all waivers by player's email 
    '****************************************************************************
    public function GetWaiversByEmail(emailaddress)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetWaiver2sByEmail"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@PassedInSiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("@PassedInEmailAddress", adVarChar, adParamInput, 200, emailaddress)
        
        
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

        GetWaiversByEmail = results

    end function

    '****************************************************************************
    ' get waiver by id 
    '****************************************************************************
    public function GetWaiversById(nWaiverId)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetWaiver2sById"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@PassedInWaiverId", adVarChar, adParamInput, 200, nWaiverId)
        
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

        GetWaiversById = results

    end function

    '****************************************************************************
    ' get waiver by hash id 
    '****************************************************************************
    public function GetWaiversByHashId(sHashId)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetWaiver2sByHashId"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@PassedInWaiverHashId", adVarChar, adParamInput, 32, sHashId)
        
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

        GetWaiversByHashId = results

    end function

    '****************************************************************************
    ' get all waivers by player's search info 
    '****************************************************************************
    public function GetWaiversByPlayersInfo(FirstName, LastName, DOB)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetWaiver2sByPlayerInfo"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@PassedInSiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
	    oCmd.Parameters.Append oCmd.CreateParameter("@PassedInFirstName", adVarChar, adParamInput, 50, FirstName)
	    oCmd.Parameters.Append oCmd.CreateParameter("@PassedInLastName", adVarChar, adParamInput, 50, LastName)
        oCmd.Parameters.Append oCmd.CreateParameter("@PassedInDateOfBirth", adDate, adParamInput, 7, DOB)
        
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

        GetWaiversByPlayersInfo = results

    end function

    '****************************************************************************
    ' return all of the waivers
    '****************************************************************************
    public function GetWaiversForExportByAll(IncludePrevious)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetWaiversForExportByAll"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("@IncludePrevious", adBoolean, adParamInput, 1, IncludePrevious)
        
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

        GetWaiversForExportByAll = results
    end function

    '****************************************************************************
    ' return the waivers created within the specified range
    '****************************************************************************
    public function GetWaiversForExportByCustomRange(StartDate, EndDate, IncludePrevious)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetWaiversForExportByCustomRange"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("@StartDate", adDate, adParamInput, 7, StartDate)
        oCmd.Parameters.Append oCmd.CreateParameter("@EndDate", adDate, adParamInput, 7, EndDate)
        oCmd.Parameters.Append oCmd.CreateParameter("@IncludePrevious", adBoolean, adParamInput, 1, IncludePrevious)
        
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

        GetWaiversForExportByCustomRange = results
    end function

    '****************************************************************************
    ' return all the waivers created within the current month
    '****************************************************************************
    public function GetWaiversForExportByMonth(IncludePrevious)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetWaiversForExportByMonth"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("@IncludePrevious", adBoolean, adParamInput, 1, IncludePrevious)
        
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

        GetWaiversForExportByMonth = results
    end function

    '****************************************************************************
    ' return all the waivers created within the current year
    '****************************************************************************
    public function GetWaiversForExportByYear(IncludePrevious)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetWaiversForExportByYear"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("@IncludePrevious", adBoolean, adParamInput, 1, IncludePrevious)
        
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

        GetWaiversForExportByYear = results
    end function

    '****************************************************************************
    ' save the waiver and return the waiver id 
    '****************************************************************************
    public function SaveWaiver(WaiverId, PlayerId, ParentId, _
                                PhoneNumber, Address, City, State, _
                                ZipCode, CreateDateTime, EmailAddress, _
                                Signature, HashId, SubmittedFromIpAddress)
        dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_SaveWaiver"
        oCmd.CommandType = adCmdStoredProc

        oCmd.Parameters.Append oCmd.CreateParameter("@returnNewWaiver2id", adInteger, adParamReturnValue)
        oCmd.Parameters.Append oCmd.CreateParameter("@WaiverId", adInteger, adParamInput, 4, WaiverId)
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("@CreateDateTime", adDate, adParamInput, 7, CreateDateTime)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayerId", adInteger, adParamInput, 4, PlayerId)
        oCmd.Parameters.Append oCmd.CreateParameter("@ParentId", adInteger, adParamInput, 4, ParentId)
        oCmd.Parameters.Append oCmd.CreateParameter("@PhoneNumber", adVarChar, adParamInput, 20, PhoneNumber)
        oCmd.Parameters.Append oCmd.CreateParameter("@Address", adVarChar, adParamInput, 100, Address)
        oCmd.Parameters.Append oCmd.CreateParameter("@City", adVarChar, adParamInput, 50, City)
        oCmd.Parameters.Append oCmd.CreateParameter("@State", adChar, adParamInput, 2, State)
        oCmd.Parameters.Append oCmd.CreateParameter("@ZipCode", adVarChar, adParamInput, 10, ZipCode)
        oCmd.Parameters.Append oCmd.CreateParameter("@EmailAddress", adVarChar, adParamInput, 200, EmailAddress)
        oCmd.Parameters.Append oCmd.CreateParameter("@LegaleseId", adInteger, adParamInput, 4, 0)
        oCmd.Parameters.Append oCmd.CreateParameter("@HashId", adVarChar, adParamInput, 32, HashId)
        oCmd.Parameters.Append oCmd.CreateParameter("@SubmittedFromIpAddress", adVarChar, adParamInput, 40, SubmittedFromIpAddress)    
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayerSignature", adVarChar, adParamInput, -1, Signature)
        oCmd.Parameters.Append oCmd.CreateParameter("@Active", adBoolean, adParamInput, , True)

        oCon.ExecuteCommand oCmd

        SaveWaiver = oCmd("@returnNewWaiver2id")

        oCon.CloseConnection
        set oCmd = nothing

    end function


    '****************************************************************************
    ' get the results from the search of the waiver table
    '****************************************************************************
    public function SearchWaivers(SearchText, IncludeInvalid, IncludeCheckedIn)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_SearchWaivers"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("@SearchText", adVarChar, adParamInput, 1000, SearchText)
        oCmd.Parameters.Append oCmd.CreateParameter("@includeInvalid", adBoolean, adParamInput, , IncludeInvalid)
        oCmd.Parameters.Append oCmd.CreateParameter("@IncludeCheckedIn", adBoolean, adParamInput, , IncludeCheckedIn)
        
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

        SearchWaivers = results
    end function

    '****************************************************************************
    ' mark all of the waivers as exported for the site
    '****************************************************************************
    public sub UpdateWaiverExportForAll()
         dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_UpdateWaiverExportForAll"
        oCmd.CommandType = adCmdStoredProc

        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)

        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub

    '****************************************************************************
    ' mark all of the waivers within the date range as exported
    '****************************************************************************
    public sub UpdateWaiverExportForCustomRange(StartDate, EndDate)
         dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_UpdateWaiverExportForCustomRange"
        oCmd.CommandType = adCmdStoredProc

        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("@StartDate", adDate, adParamInput, 7, StartDate)
        oCmd.Parameters.Append oCmd.CreateParameter("@EndDate", adDate, adParamInput, 7, EndDate)
        
        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub

    '****************************************************************************
    ' mark all of the waivers for the month as exported
    '****************************************************************************
    public sub UpdateWaiverExportForMonth()
         dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_UpdateWaiverExportForMonth"
        oCmd.CommandType = adCmdStoredProc

        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)

        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub

    '****************************************************************************
    ' mark all of the waivers for the year as exported
    '****************************************************************************
    public sub UpdateWaiverExportForYear()
         dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_UpdateWaiverExportForYear"
        oCmd.CommandType = adCmdStoredProc

        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)

        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub
end class
%>
