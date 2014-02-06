<%
class GlobalStatistics
    public function GetTotalGuestCount()
        dim gsCon : set gsCon = new GlobalStatisticsDataConnection
        dim results : results = gsCon.GetTotalGuestCount()

        if ubound(results) >= 0 then
            GetTotalGuestCount = results(0,0)
        else
            GetTotalGuestCount = 0
        end if
    end function

    public function GetTotalScheduledEventCount()
        dim gsCon : set gsCon = new GlobalStatisticsDataConnection
        dim results : results = gsCon.GetTotalScheduledEventCount()

        if ubound(results) >= 0 then
            GetTotalScheduledEventCount = results(0,0)
        else
            GetTotalScheduledEventCount = 0
        end if
    end function

    public function GetTotalWaiverCount()
        dim gsCon : set gsCon = new GlobalStatisticsDataConnection
        dim results : results = gsCon.GetTotalWaiverCount()

        if ubound(results) >= 0 then
            GetTotalWaiverCount = results(0,0)
        else
            GetTotalWaiverCount = 0
        end if
    end function
end class
%>