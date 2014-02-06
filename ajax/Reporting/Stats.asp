<!--#include file="../../classes/IncludeList.asp"-->
<%
    const AJAXACTION_STATS_ALL = "0"

    select case request.QueryString(QUERYSTRING_VAR_ACTION)
        case AJAXACTION_STATS_ALL
            BuildGlobalStatisticsJSON
    end select

    private sub BuildGlobalStatisticsJSON()
        dim statDict : set statDict = Server.CreateObject("Scripting.Dictionary")
        dim gs : set gs = new GlobalStatistics

        statDict.Add "Events", FormatNumber(gs.GetTotalScheduledEventCount(), 0)
        statDict.Add "Guests", FormatNumber(gs.GetTotalGuestCount(), 0)
        statDict.Add "Waivers", FormatNumber(gs.GetTotalWaiverCount(), 0)

        response.Clear()
        response.ContentType = MIMETYPE_JSON

        dim jsonParser
        set jsonParser = new JSON
        response.Write jsonParser.toJSON_v2(statDict)
    end sub
%>