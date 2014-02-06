<%
class PlayerDataConnection

    '****************************************************************************
    ' get player by id 
    '****************************************************************************
    public function GetPlayerById(PlayerId)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetPlayerById"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@PassedInPlayerId", adInteger, adParamInput, 4, PlayerId)
        
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

        GetPlayerById = results

    end function

    '****************************************************************************
    ' get players ids by name dob 
    '****************************************************************************
    function GetPlayersByNameDOB(FirstName, LastName, DOB)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetPlayersByNameDOB"
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

        GetPlayersByNameDOB = results

    end function

    '****************************************************************************
    ' get players by name dob address
    '****************************************************************************
    function GetPlayersByNameDOBAddress(FirstName, LastName, DOB, Addy)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetPlayersByNameDOBAddress"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@PassedInSiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
	    oCmd.Parameters.Append oCmd.CreateParameter("@PassedInFirstName", adVarChar, adParamInput, 50, FirstName)
	    oCmd.Parameters.Append oCmd.CreateParameter("@PassedInLastName", adVarChar, adParamInput, 50, LastName)
        oCmd.Parameters.Append oCmd.CreateParameter("@PassedInDateOfBirth", adDate, adParamInput, 7, DOB)
	    oCmd.Parameters.Append oCmd.CreateParameter("@PassedInAddress", adVarChar, adParamInput, 100, Addy)
        
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

        GetPlayersByNameDOBAddress = results

    end function

    '****************************************************************************
    ' get player by info 
    '****************************************************************************
    public function GetPlayerBySearchInfo(FirstName, LastName, DOB)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetPlayerIdForWaiverSearch"
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

    for each x in results
    response.write "@ " & x
    next

        GetPlayerBySearchInfo = results

    end function

    '****************************************************************************
    ' get random streets 
    '****************************************************************************
    public function GetRandomStreets()
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetRandomStreets"
        oCmd.CommandType = adCmdStoredProc
        
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

        GetRandomStreets = results

    end function

    '****************************************************************************
    ' save the player and return the player id 
    '****************************************************************************
    public function SavePlayer(PlayerId, EmailList, FirstName, LastName, _
                                    DateOfBirth, PhoneNumber, _
                                    Address, City, State, _
                                    ZipCode, EmailAddress)
        dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_SavePlayer"
        oCmd.CommandType = adCmdStoredProc

        oCmd.Parameters.Append oCmd.CreateParameter("@returnPlayerid", adInteger, adParamReturnValue)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayerId", adInteger, adParamInput, 4, PlayerId)
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)

        if IsNull(FirstName) = false and len(FirstName) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@FirstName", adVarChar, adParamInput, 50, FirstName)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@FirstName", adVarChar, adParamInput, 50, null)
        end if
        if IsNull(LastName) = false and len(LastName) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@LastName", adVarChar, adParamInput, 50, LastName)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@LastName", adVarChar, adParamInput, 50, null)
        end if
        if IsNull(DateOfBirth) = false and len(DateOfBirth) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@DateOfBirth", adDate, adParamInput, 7, DateOfBirth)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@DateOfBirth", adDate, adParamInput, 7, null)
        end if
        if IsNull(PhoneNumber) = false and len(PhoneNumber) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@PhoneNumber", adVarChar, adParamInput, 20, PhoneNumber)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@PhoneNumber", adVarChar, adParamInput, 20, null)
        end if
        if IsNull(Address) = false and len(Address) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@Address", adVarChar, adParamInput, 100, Address)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@Address", adVarChar, adParamInput, 100, null)
        end if
        if IsNull(City) = false and len(City) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@City", adVarChar, adParamInput, 50, City)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@City", adVarChar, adParamInput, 50, null)
        end if
        if IsNull(State) = false and len(State) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@State", adChar, adParamInput, 2, State)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@State", adChar, adParamInput, 2, null)
        end if
        if IsNull(ZipCode) = false and len(ZipCode) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@ZipCode", adVarChar, adParamInput, 10, ZipCode)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@ZipCode", adVarChar, adParamInput, 10, null)
        end if
        if IsNull(EmailAddress) = false and len(EmailAddress) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@EmailAddress", adVarChar, adParamInput, 200, EmailAddress)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@EmailAddress", adVarChar, adParamInput, 200, null)
        end if
        if IsNull(EmailList) = false and len(EmailList) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@EmailList", adBoolean, adParamInput, 1, EmailList)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@EmailList", adBoolean, adParamInput, 1, null)
        end if
        
        oCon.ExecuteCommand oCmd

        SavePlayer = oCmd("@returnPlayerid")

        oCon.CloseConnection
        set oCmd = nothing
    end function

end class
%>
