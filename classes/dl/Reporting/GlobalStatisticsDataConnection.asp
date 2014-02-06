<%
class GlobalStatisticsDataConnection
    public function GetTotalGuestCount()
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.RecordSet")
        set oCon = new DataConnection

        oCmd.CommandText = "GetStatisticTotalGuests"
        oCmd.CommandType = adCmdStoredProc

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
    
        GetTotalGuestCount = results
    end function

    public function GetTotalScheduledEventCount()
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.RecordSet")
        set oCon = new DataConnection

        oCmd.CommandText = "GetStatisticTotalScheduledEvents"
        oCmd.CommandType = adCmdStoredProc

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

        GetTotalScheduledEventCount = results
    end function

    public function GetTotalWaiverCount()
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.RecordSet")
        set oCon = new DataConnection

        oCmd.CommandText = "GetStatisticTotalWaivers"
        oCmd.CommandType = adCmdStoredProc

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

        GetTotalWaiverCount = results
    end function
end class
%>