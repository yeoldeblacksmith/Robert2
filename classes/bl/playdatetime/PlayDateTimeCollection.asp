<%
class PlayDateTimeCollection
    private mo_List
    
    'ctor
    public sub Class_Initialize
        set mo_List = new ArrayList
    end sub

    ' methods
    public sub Add(value)
        mo_List.Add value
    end sub

    public sub LoadAllByPlayer(nPlayerId)
        dim myCon
        set myCon = new PlayDateTimeDataConnection

        dim PlayDateTimeColl
        PlayDateTimeColl = myCon.GetAllPlayDateTimeByPlayerId(nPlayerId)

        if UBound(PlayDateTimeColl) > 0 then
            for i = 0 to ubound(PlayDateTimeColl,2)
                dim myPlayDate
                set myPlayDate = new PlayDateTime

                myPlayDate.PlayerId = PlayDateTimeColl(PLAYDATETIME_INDEX_PLAYERID, i)
                myPlayDate.PlayDate = PlayDateTimeColl(PLAYDATETIME_INDEX_PLAYDATE, i)
                myPlayDate.PlayTime = PlayDateTimeColl(PLAYDATETIME_INDEX_PLAYTIME, i)
                myPlayDate.EventId = PlayDateTimeColl(PLAYDATETIME_INDEX_EVENTID, i)
                myPlayDate.CheckInDateTimeAdmin = PlayDateTimeColl(PLAYDATETIME_INDEX_CHECKINDATETIMEADMIN, i)

                mo_List.Add myPlayDate
            next
        end if
    end sub

    public sub LoadAllByPlayerAndDate(nPlayerId,dPlayDate)
        dim myCon
        set myCon = new PlayDateTimeDataConnection

        dim PlayDateTimeColl
        PlayDateTimeColl = myCon.GetAllPlayDateTimeByPlayerIdAndDate(nPlayerId, dPlayDate)

        if UBound(PlayDateTimeColl) > 0 then
            for i = 0 to ubound(PlayDateTimeColl,2)
                dim myPlayDate
                set myPlayDate = new PlayDateTime

                myPlayDate.PlayerId = PlayDateTimeColl(PLAYDATETIME_INDEX_PLAYERID, i)
                myPlayDate.PlayDate = PlayDateTimeColl(PLAYDATETIME_INDEX_PLAYDATE, i)
                myPlayDate.PlayTime = PlayDateTimeColl(PLAYDATETIME_INDEX_PLAYTIME, i)
                myPlayDate.EventId = PlayDateTimeColl(PLAYDATETIME_INDEX_EVENTID, i)
                myPlayDate.CheckInDateTimeAdmin = PlayDateTimeColl(PLAYDATETIME_INDEX_CHECKINDATETIMEADMIN, i)

                mo_List.Add myPlayDate
            next
        end if
    end sub

    public sub Remove(value)
        mo_List.Remove value
    end sub

    public function ToJSON()
        dim output

        output = "["
    
        for i = 0 to Count - 1
            if i > 0 then output = output & ","
            output = output & Item(i).ToJSON()
        next    

        output = output & "]"

        ToJSON = output
    end function


    'properties
    public property get Count
        Count = mo_List.Count
    end property

    Public Default Property Get Item(index)
        set Item = mo_List(index)
    end property
end class
%>