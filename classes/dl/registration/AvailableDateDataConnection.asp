<%
class AvailableDateDataConnection
    sub DeleteAvailableDate(SelectedDate)
        dim oCmd, oCon
        set oCmd = Server.CreateObject("ADODB.Command")

        oCmd.CommandText = "DeleteAvailableDate"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("AvailableDate", adDBDate, adParamInput, 8, SelectedDate)
        
        set oCon = new DataConnection
        oCon.ExecuteCommand oCmd

        set oCmd = nothing
        oCon.CloseConnection
    end sub

    function GetAvailableDate(SelectedDate)
        dim oCmd, oRs, oCon
        set oCmd = Server.CreateObject("ADODB.Command")

        oCmd.CommandText = "GetAvailableDate"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("AvailableDate", adDBDate, adParamInput, 8, SelectedDate)

        set oCon = new DataConnection
        set oRs = Server.CreateObject("ADODB.Recordset")
        oCon.GetRecordset oCmd, oRs

        if oRs.EOF = false then
            dim results(3)
    
            results(AVAILABLEDATE_INDEX_DATE) = oRs(AVAILABLEDATE_COLUMN_DATE)
            results(AVAILABLEDATE_INDEX_STARTTIME) = Left(oRs(AVAILABLEDATE_COLUMN_STARTTIME),5)
            results(AVAILABLEDATE_INDEX_ENDTIME) = Left(oRs(AVAILABLEDATE_COLUMN_ENDTIME),5)

            GetAvailableDate = results
        end if

        oRs.Close

        set oRs = nothing
        set oCmd = nothing

        oCon.CloseConnection
    end function

    sub SaveAvailableDate(SelectedDate, StartTime, EndTime)
        dim oCmd, oCon
        set oCmd = Server.CreateObject("ADODB.Command")

        oCmd.CommandText = "SaveAvailableDate"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("AvailableDate", adDBDate, adParamInput, 8, SelectedDate)
        oCmd.Parameters.Append oCmd.CreateParameter("StartTime", adDBTime, adParamInput, 6, StartTime)
        oCmd.Parameters.Append oCmd.CreateParameter("EndTime", adDBTime, adParamInput, 6, EndTime)
        
        set oCon = new DataConnection
        oCon.ExecuteCommand oCmd
        oCon.CloseConnection

        set oCmd = nothing
    end sub
end class
%>