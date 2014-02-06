<%
class SiteSetting
    private ms_Key, ms_Value

    public sub Load(sKey)
        if len(sKey) = 0 then exit sub

        dim myCon
        set myCon = new SiteSettingsDataConnection

        dim myResults
        myResults = myCon.GetSiteSettings(MY_SITE_GUID, sKey)

        if UBound(myResults) > 0 then
            'SiteGuid = Results(SITESETTINGS_INDEX_SITEGUID, 0)
            Key = Results(SITESETTINGS_INDEX_KEY, 0)
            Value = Results(SITESETTINGS_INDEX_KEY, 0)
        end if
    end sub

    public sub Save()
        dim myCon
        set myCon = new SiteSettingsDataConnection

        myCon.SaveSiteSetting MY_SITE_GUID, Key, Value 
    end sub

    public property get Key
        Key = ms_Key
    end property

    public property let Key(value)
        ms_Key = value
    end property

    public property get Value
        Value = ms_Value
    end property

    public property let Value(val)
        ms_Value = val
    end property
end class
%>