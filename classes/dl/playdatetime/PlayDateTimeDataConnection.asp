<%
class PlayDateTimeDataConnection

    '****************************************************************************
    ' Add New PlayDateTime Record 
    '****************************************************************************
    public sub AddNewPlayDateTime(PlayerId, PlayDate, PlayTime, EventId, CheckInDateTimeAdmin)
        dim oCmd, oCon
    
        If EventId = "" Then EventId = NULL End If
        'PlayTime = FormatDateTime(PlayTime,4)
        PlayTime = FormatDateTime(PlayTime,4)
 
        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_AddNewPlayDateTime"
        oCmd.CommandType = adCmdStoredProc

        oCmd.Parameters.Append oCmd.CreateParameter("@PlayerId", adInteger, adParamInput, 4, PlayerId)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayDate", adDate, adParamInput, 7, PlayDate)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayTime", adDBTime, adParamInput, 0, PlayTime)
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)
        oCmd.Parameters.Append oCmd.CreateParameter("@CheckInDateTimeAdmin", adDBTimeStamp, adParamInput, , CheckInDateTimeAdmin)

        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub
 
    '****************************************************************************
    ' Delete PlayDateTime Record 
    '****************************************************************************
    public sub Delete(PlayerId, PlayDate, PlayTime)
        dim oCmd, oCon

        PlayTime = Left(PlayTime,8)
    
        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_DeletePlayDateTime"
        oCmd.CommandType = adCmdStoredProc

        oCmd.Parameters.Append oCmd.CreateParameter("@PlayerId", adInteger, adParamInput, 4, PlayerId)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayDate", adDate, adParamInput, 7, PlayDate)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayTime", adDBTime, adParamInput, 0, PlayTime)

        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub

    '****************************************************************************
    ' save the playdatetime 
    '****************************************************************************
    public sub SavePlayDateTime(PlayerId, PlayDate, PlayTime, EventId, CheckInDateTimeAdmin)
        dim oCmd, oCon
    
        If EventId = "" Then EventId = NULL End If
        PlayTime = Left(PlayTime,8)

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_SavePlayDateTime"
        oCmd.CommandType = adCmdStoredProc

        oCmd.Parameters.Append oCmd.CreateParameter("@PlayerId", adInteger, adParamInput, 4, PlayerId)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayDate", adDate, adParamInput, 7, PlayDate)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayTime", adDBTime, adParamInput, 0, PlayTime)
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)
        oCmd.Parameters.Append oCmd.CreateParameter("@CheckInDateTimeAdmin", adDBTimeStamp, adParamInput, , CheckInDateTimeAdmin)

        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub

    '****************************************************************************
    ' update existing playdatetime record 
    '****************************************************************************
    public sub UpdatePlayDateTime(PlayerId, PlayDate, PlayTime, EventId, CheckInDateTimeAdmin, originalPlayDate, originalPlayTime)
        dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_ChangePlayDateTime"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayerId", adInteger, adParamInput, 4, PlayerId)
        oCmd.Parameters.Append oCmd.CreateParameter("@originalPlayDate", adDate, adParamInput, 7, originalPlayDate)
        oCmd.Parameters.Append oCmd.CreateParameter("@originalPlayTime", adDBTime, adParamInput, 0, originalPlayTime)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayDate", adDate, adParamInput, 7, PlayDate)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayTime", adDBTime, adParamInput, 0, PlayTime)
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)
        oCmd.Parameters.Append oCmd.CreateParameter("@CheckInDateTimeAdmin", adDBTimeStamp, adParamInput, , CheckInDateTimeAdmin)

        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub

    '****************************************************************************
    ' update existing playdatetime record - change event id and playtime
    '****************************************************************************
    public sub UpdatePlayDateEventIdAndTime(PlayerId, PlayDate, PlayTime, EventId, CheckInDateTimeAdmin, oldPlayTime)
        dim oCmd, oCon
    
        If EventId = "" Then EventId = NULL End If
        PlayTime = Left(PlayTime,8)
        oldPlayTime = Left(oldPlayTime,8)

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_ChangePlayDateTimeEventAndTime"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayerId", adInteger, adParamInput, 4, PlayerId)
        oCmd.Parameters.Append oCmd.CreateParameter("@oldPlayTime", adDBTime, adParamInput, 0, oldPlayTime)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayDate", adDate, adParamInput, 7, PlayDate)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayTime", adDBTime, adParamInput, 0, PlayTime)
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)
        oCmd.Parameters.Append oCmd.CreateParameter("@CheckInDateTimeAdmin", adDBTimeStamp, adParamInput, , CheckInDateTimeAdmin)

        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub

    '****************************************************************************
    ' update existing playdatetime record - change date
    '****************************************************************************
    public sub UpdatePlayDateTimeDate(PlayerId, PlayDate, PlayTime, EventId, CheckInDateTimeAdmin, oldPlayDate)
        dim oCmd, oCon
    
        If EventId = "" Then EventId = NULL End If
        PlayTime = Left(PlayTime,8)

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_ChangePlayDateTimeDate"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayerId", adInteger, adParamInput, 4, PlayerId)
        oCmd.Parameters.Append oCmd.CreateParameter("@oldPlayDate", adDate, adParamInput, 7, oldPlayDate)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayDate", adDate, adParamInput, 7, PlayDate)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayTime", adDBTime, adParamInput, 0, PlayTime)
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)
        oCmd.Parameters.Append oCmd.CreateParameter("@CheckInDateTimeAdmin", adDBTimeStamp, adParamInput, , CheckInDateTimeAdmin)

        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub

    '****************************************************************************
    ' update playdatetime checkinadmin 
    '****************************************************************************
    public sub UpdatePlayDateTimeCheckInAdmin(PlayerId, PlayDate, PlayTime, CheckInDateTimeAdmin)
        dim oCmd, oCon
    
        If EventId = "" Then EventId = NULL End If
        PlayTime = Left(PlayTime,5)

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_UpdatePlayDateTimeCheckInAdmin"
        oCmd.CommandType = adCmdStoredProc

        oCmd.Parameters.Append oCmd.CreateParameter("@PlayerId", adInteger, adParamInput, 4, PlayerId)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayDate", adDate, adParamInput, 7, PlayDate)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayTime", adDBTime, adParamInput, 0, PlayTime)
        oCmd.Parameters.Append oCmd.CreateParameter("@CheckInDateTimeAdmin", adDBTimeStamp, adParamInput, , CheckInDateTimeAdmin)

        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub

    '****************************************************************************
    ' Get all playdatetime records by player id
    '****************************************************************************
    public function GetAllPlayDateTimeByPlayerId(PlayerId)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetAllPlayDateTimeByPlayerId"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayerId", adInteger, adParamInput, 4, PlayerId)

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

        GetAllPlayDateTimeByPlayerId = results
    end function

    '****************************************************************************
    ' Get all playdatetime records by player id and date
    '****************************************************************************
    public function GetAllPlayDateTimeByPlayerIdAndDate(PlayerId, PlayDate)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetPlayDateTimeByPlayerAndDate"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayerId", adInteger, adParamInput, 4, PlayerId)
            oCmd.Parameters.Append oCmd.CreateParameter("@PlayDate", adDate, adParamInput, 7, PlayDate)


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

        GetAllPlayDateTimeByPlayerIdAndDate = results
    end function

    '****************************************************************************
    ' Get playdatetime record
    '****************************************************************************
    public function GetPlayDateTime(PlayerId, PlayDate, PlayTime)
        dim oCmd, oRs, oCon

        PlayTime = Left(PlayTime,8)

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetPlayDateTime"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayerId", adInteger, adParamInput, 4, PlayerId)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayDate", adDate, adParamInput, 7, PlayDate)
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayTime", adDBTime, adParamInput, 0, PlayTime)

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

        GetPlayDateTime = results
    end function

    '****************************************************************************
    ' Get all playdatetime records by player id
    '****************************************************************************
    public function GetTodaysCheckInStatus(PlayerId)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "W2_GetTodaysCheckInStatus"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@PlayerId", adInteger, adParamInput, 4, PlayerId)

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

        GetTodaysCheckInStatus = results
    end function

end class
%>
