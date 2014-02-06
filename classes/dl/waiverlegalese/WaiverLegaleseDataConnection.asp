<%
class WaiverLegaleseDataConnection

    public function GetCurrentWaiverLegaleseBySite()
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetCurrentWaiverLegaleseBySite"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        
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

        GetCurrentWaiverLegaleseBySite = results
    end function

    public function GetWaiverLegalese(LegaleseId)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetWaiverLegalese"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@LegaleseId", adInteger, adParamInput, 4, LegaleseId)
        
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

        GetWaiverLegalese = results
    end function
end class
%>