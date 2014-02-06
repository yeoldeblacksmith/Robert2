 <%
class PlayDateTime
    ' attributes
    dim mb_Loaded, mn_PlayerId, mdt_PlayDate, mt_PlayTime, mn_EventId, mdt_CheckInDateTimePlayer, mdt_CheckInDateTimeAdmin

    ' ctor
    public sub Class_Initialize
        mb_Loaded = false
    end sub

    ' methods
    public sub AddNew()
        dim myCon
        set myCon = new PlayDateTimeDataConnection

        myCon.AddNewPlayDateTime PlayerId, PlayDate, PlayTime, EventId, CheckInDateTimeAdmin
    end sub

    public sub Delete()
        dim myCon
        set myCon = new PlayDateTimeDataConnection

        myCon.Delete PlayerId, PlayDate, PlayTime
    end sub

    public sub GetTodaysCheckInStatus(nPlayerId)
        dim myCon
        set myCon = new PlayDateTimeDataConnection

        dim results
        results = myCon.GetTodaysCheckInStatus(nPlayerId)

        if UBound(results) > 0 then
            PlayerId = results(PLAYDATETIME_INDEX_PLAYERID, 0)
            PlayDate = results(PLAYDATETIME_INDEX_PLAYDATE, 0)
            PlayTime = results(PLAYDATETIME_INDEX_PLAYTIME, 0)
            EventId = results(PLAYDATETIME_INDEX_EVENTID, 0)
            CheckInDateTimeAdmin = results(PLAYDATETIME_INDEX_CHECKINDATETIMEADMIN, 0)

            mb_Loaded = true
        end if       
    end sub


    public sub Load(nPlayerId, dt_PlayDate, t_PlayTime)
        dim myCon
        set myCon = new PlayDateTimeDataConnection

        dim results
        results = myCon.GetPlayDateTime(nPlayerId, dt_PlayDate, t_PlayTime)

        if UBound(results) > 0 then
            PlayerId = results(PLAYDATETIME_INDEX_PLAYERID, 0)
            PlayDate = results(PLAYDATETIME_INDEX_PLAYDATE, 0)
            PlayTime = results(PLAYDATETIME_INDEX_PLAYTIME, 0)
            EventId = results(PLAYDATETIME_INDEX_EVENTID, 0)
            CheckInDateTimeAdmin = results(PLAYDATETIME_INDEX_CHECKINDATETIMEADMIN, 0)

            mb_Loaded = true
        end if
    end sub
     
    public sub Save()
        dim myCon
        set myCon = new PlayDateTimeDataConnection

        myCon.SavePlayDateTime PlayerId, PlayDate, PlayTime, EventId, CheckInDateTimeAdmin
    end sub

    public function ToJSON()
        ToJSON = (new JSON).toJSON_V2(Me)
    end function

    public sub UpdateAdminCheckIn()
        dim myCon
        set myCon = new PlayDateTimeDataConnection
     
        myCon.UpdatePlayDateTimeCheckInAdmin PlayerId, PlayDate, PlayTime, CheckInDateTimeAdmin
    end sub

    public sub UpdatePlayDateTime(originalPlayDate, originalPlayTime)
        dim myCon
        set myCon = new PlayDateTimeDataConnection

        myCon.UpdatePlayDateTime PlayerId, PlayDate, PlayTime, EventId, CheckInDateTimeAdmin, originalPlayDate, originalPlayTime
    end sub


    public sub UpdateDate(oldPlayDate)
        dim myCon
        set myCon = new PlayDateTimeDataConnection

        myCon.UpdatePlayDateTimeDate PlayerId, PlayDate, PlayTime, EventId, CheckInDateTimeAdmin, oldPlayDate
    end sub

    public sub UpdatePlayDateEventAndTime(oldPlayTime)
        dim myCon
        set myCon = new PlayDateTimeDataConnection

        myCon.UpdatePlayDateEventIdAndTime PlayerId, PlayDate, PlayTime, EventId, CheckInDateTimeAdmin, oldPlayTime
    end sub

    ' properties

    public property get CheckInDateTimeAdmin
        CheckInDateTimeAdmin = mdt_CheckInDateTimeAdmin
    end property

    public property let CheckInDateTimeAdmin(value)
        mdt_CheckInDateTimeAdmin = value
    end property

    public property get EventId
        EventId = mn_EventId
    end property

    public property let EventId(value)
        mn_EventId = value
    end property

    public property get Loaded
        Loaded = mb_Loaded
    end property

    public property get PlayDate
        PlayDate = mdt_PlayDate
    end property

    public property let PlayDate(value)
        mdt_PlayDate = value
    end property

    public property get PlayerId
        PlayerId = mn_PlayerId
    end property

    public property let PlayerId(value)
        mn_PlayerId = value
    end property

    public property get PlayTime
        PlayTime = mt_PlayTime
    end property

    public property let PlayTime(value)
        mt_PlayTime = value
    end property

    public property get PartyName
        Dim se 
        Set se = new ScheduledEvent
        se.Load EventId
        PartyName = se.PartyName
    end property

end class
%>