<% 
class SiteSettingsDataConnection
    function GetAllSiteSettingsBySite(SiteGuid)
        dim oCmd, oRs, oCon
        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetAllSiteSettingsBySite"
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

        GetAllSiteSettingsBySite = results
    end function

    function GetSiteSettings(SiteGuid, Key)
        dim oCmd, oRs, oCon
        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetSiteSettings"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, SiteGuid)
        oCmd.Parameters.Append oCmd.CreateParameter("@Key", adVarChar, adParamInput, 128, Key)

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

        GetSiteSettings = results
    end function

    sub SaveSiteSetting(SiteGuid, Key, Value)

        dim oCmd, oCon
        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "SaveSiteSetting"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, SiteGuid)
        oCmd.Parameters.Append oCmd.CreateParameter("@Key", adVarChar, adParamInput, 128, Key)

        if len(Value) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@Value", adVarChar, adParamInput, -1, Value)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@Value", adVarChar, adParamInput, -1, Null)
        end if
    
        oCon.ExecuteCommand oCmd
        oCon.CloseConnection

        set oCmd = nothing
    end sub
end class
%>