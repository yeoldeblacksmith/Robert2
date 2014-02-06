<%
class WaiverLegalese
    ' attributes
    private ms_SiteGuid, mn_LegaleseId, mb_Active, mdt_CreatedDateTime, mdt_PostedDateTime, ms_LegaleseText, mb_Loaded

    ' ctor
    public sub Class_Initialize
        mb_Loaded = false
    end sub

    ' methods
    public sub LoadCurrentBySite()
        dim myCon
        set myCon = new WaiverLegaleseDataConnection

        dim results
        results = myCon.GetCurrentWaiverLegaleseBySite()

        if UBound(results) > 0 then
            LegaleseId = results(WAIVERLEGALESE_INDEX_LEGALESEID,0)
            Active = results(WAIVERLEGALESE_INDEX_ACTIVE,0)
            SiteGuid = results(WAIVERLEGALESE_INDEX_SITEGUID,0)
            CreatedDateTime = results(WAIVERLEGALESE_INDEX_CREATEDDATETIME,0)
            PostedDateTime = results(WAIVERLEGALESE_INDEX_POSTEDDATETIME,0)
            LegaleseText = results(WAIVERLEGALESE_INDEX_LEGALESETEXT,0)

            mb_Loaded = true
        end if
    end sub

    public sub Load(nLegaleseId)
        dim myCon
        set myCon = new WaiverLegaleseDataConnection

        dim results
        results = myCon.GetWaiverLegalese(nLegaleseId)

        if UBound(results) > 0 then
            LegaleseId = results(WAIVERLEGALESE_INDEX_LEGALESEID,0)
            Active = results(WAIVERLEGALESE_INDEX_ACTIVE,0)
            SiteGuid = results(WAIVERLEGALESE_INDEX_SITEGUID,0)
            CreatedDateTime = results(WAIVERLEGALESE_INDEX_CREATEDDATETIME,0)
            PostedDateTime = results(WAIVERLEGALESE_INDEX_POSTEDDATETIME,0)
            LegaleseText = results(WAIVERLEGALESE_INDEX_LEGALESETEXT,0)

            mb_Loaded = true
        end if
    end sub

    'properties

    public property get LegaleseId
        LegaleseId = mn_LegaleseId
    end property

    public property let LegaleseId(value)
        mn_LegaleseId = value
    end property

    public property get Active
        Active  = mb_Active
    end property

    public property let Active(value)
        mb_Active = value
    end property
    
    public property get SiteGuid
        SiteGuid = ms_SiteGuid
    end property

    public property let SiteGuid(value)
        ms_SiteGuid = value
    end property

    public property get CreatedDateTime
        CreatedDateTime = mdt_CreatedDateTime
    end property

    public property let CreatedDateTime(value)
        mdt_CreatedDateTime = value
    end property
    
    public property get PostedDateTime
        PostedDateTime = mdt_PostedDateTime
    end property

    public property let PostedDateTime(value)
        mdt_PostedDateTime = value
    end property

    public property get LegaleseText
        LegaleseText = ms_LegaleseText
    end property

    public property let LegaleseText(value)
        ms_LegaleseText = value
    end property

end class
%>