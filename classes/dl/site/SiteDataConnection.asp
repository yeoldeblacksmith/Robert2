<%
class SiteDataConnection
    function GetSite(SiteGuid)
        dim oCmd, oRs, oCon
        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetSite"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("SiteGuid", adVarChar, adParamInput, 38, SiteGuid)

        oCon.GetRecordSet oCmd, oRs

        dim results
        if oRs.EOF then
            results = Array()
        else
            results = oRs.GetRows()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetSite = results
    end function

    function GetSiteByName(SiteName)
        dim oCmd, oRs, oCon
        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetSiteByName"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@Name", adVarChar, adParamInput, 200, SiteName)

        oCon.GetRecordSet oCmd, oRs

        dim results
        if oRs.EOF then
            results = Array()
        else
            results = oRs.GetRows()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetSiteByName = results
    end function

    sub SaveSite(SiteGuid, Name, AdminEmail, FromEmailAddress, FromName, MondayOpenTime, MondayCloseTime, _
                 TuesdayOpenTime, TuesdayCloseTime, WednesdayOpenTime, WednesdayCloseTime, _
                 ThursdayOpenTime, ThursdayCloseTime, FridayOpenTime, FridayCloseTime, SaturdayOpenTime, _
                 SaturdayCloseTime, SundayOpenTime, SundayCloseTime, HomeUrl, VantoraUrl, Address, City, _
                 State, ZipCode, PhoneNumber, DirectoryName, Country)

        dim oCmd, oCon
        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "SaveSite"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, SiteGuid)
        if len(Name) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@Name", adVarChar, adParamInput, 200, Name)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@Name", adVarChar, adParamInput, 200, Null)
        end if
        if len(AdminEmail) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@AdminEmail", adVarChar, adParamInput, 128, AdminEmail)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@AdminEmail", adVarChar, adParamInput, 128, Null)
        end if
        if len(FromEmailAddress) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@FromEmailAddress", adVarChar, adParamInput, 128, FromEmailAddress )
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@FromEmailAddress", adVarChar, adParamInput, 128, Null)
        end if
        if len(FromName) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@FromName", adVarChar, adParamInput, 38, FromName)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@FromName", adVarChar, adParamInput, 38, Null)
        end if
        if len(MondayOpenTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@MondayOpenTime", adDBTime, adParamInput, 6, MondayOpenTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@MondayOpenTime", adDBTime, adParamInput, 6, null)
        end if
        if len(MondayCloseTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@MondayCloseTime", adDBTime, adParamInput, 6, MondayCloseTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@MondayCloseTime", adDBTime, adParamInput, 6, null)
        end if
        if len(TuesdayOpenTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@TuesdayOpenTime", adDBTime, adParamInput, 6, TuesdayOpenTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@TuesdayOpenTime", adDBTime, adParamInput, 6, null)
        end if
        if len(TuesdayCloseTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@TuesdayCloseTime", adDBTime, adParamInput, 6, TuesdayCloseTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@TuesdayCloseTime", adDBTime, adParamInput, 6, null)
        end if
        if len(WednesdayOpenTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@WednesdayOpenTime", adDBTime, adParamInput, 6, WednesdayOpenTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@WednesdayOpenTime", adDBTime, adParamInput, 6, null)
        end if
        if len(WednesdayCloseTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@WednesdayCloseTime", adDBTime, adParamInput, 6, WednesdayCloseTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@WednesdayCloseTime", adDBTime, adParamInput, 6, null)
        end if
        if len(ThursdayOpenTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@ThursdayOpenTime", adDBTime, adParamInput, 6, ThursdayOpenTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@ThursdayOpenTime", adDBTime, adParamInput, 6, null)
        end if
        if len(ThursdayCloseTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@ThursdayCloseTime", adDBTime, adParamInput, 6, ThursdayCloseTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@ThursdayCloseTime", adDBTime, adParamInput, 6, null)
        end if
        if len(FridayOpenTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@FridayOpenTime", adDBTime, adParamInput, 6, FridayOpenTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@FridayOpenTime", adDBTime, adParamInput, 6, null)
        end if
        if len(FridayCloseTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@FridayCloseTime", adDBTime, adParamInput, 6, FridayCloseTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@FridayCloseTime", adDBTime, adParamInput, 6, null)
        end if
        if len(SaturdayOpenTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@SaturdayOpenTime", adDBTime, adParamInput, 6, SaturdayOpenTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@SaturdayOpenTime", adDBTime, adParamInput, 6, null)
        end if
        if len(SaturdayCloseTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@SaturdayCloseTime", adDBTime, adParamInput, 6, SaturdayCloseTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@SaturdayCloseTime", adDBTime, adParamInput, 6, null)
        end if
        if len(SundayOpenTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@SundayOpenTime", adDBTime, adParamInput, 6, SundayOpenTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@SundayOpenTime", adDBTime, adParamInput, 6, null)
        end if
        if len(SundayCloseTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@SundayCloseTime", adDBTime, adParamInput, 6, SundayCloseTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@SundayCloseTime", adDBTime, adParamInput, 6, null)
        end if
        if len(HomeUrl) = 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@HomeUrl", adVarChar, adParamInput, 256, null)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@HomeUrl", adVarChar, adParamInput, 256, HomeUrl)
        end if
        if len(VantoraUrl) = 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@VantoraUrl", adVarChar, adParamInput, 256, null)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@VantoraUrl", adVarChar, adParamInput, 256, VantoraUrl)
        end if
        if len(Address) = 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@Address", adVarChar, adParamInput, 256, null)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@Address", adVarChar, adParamInput, 256, Address)
        end if
        if len(City) = 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@City", adVarChar, adParamInput, 50, null)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@City", adVarChar, adParamInput, 50, City)
        end if
        if len(State) = 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@State", adChar, adParamInput, 2, null)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@State", adChar, adParamInput, 2, State)
        end if
        if len(ZipCode) = 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@ZipCode", adVarChar, adParamInput, 20, null)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@ZipCode", adVarChar, adParamInput, 20, ZipCode)
        end if
        if len(PhoneNumber) = 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@PhoneNumber", adVarChar, adParamInput, 20, null)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@PhoneNumber", adVarChar, adParamInput, 20, PhoneNumber)
        end if
        if len(DirectoryName) = 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@DirectoryName", adVarChar, adParamInput, 100, null)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@DirectoryName", adVarChar, adParamInput, 100, DirectoryName)
        end if
        if len(Country) = 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@Country", adVarChar, adParamInput, 2, null)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@Country", adVarChar, adParamInput, 2, Country)
        end if

        oCon.ExecuteCommand oCmd
        oCon.CloseConnection

        set oCmd = nothing
    end sub
end class
%>