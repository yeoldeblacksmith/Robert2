<%
class SiteSettingCollection
    ' attributes
    private mo_List

    ' ctor
    public sub Class_Initialize()
        set mo_List = new ArrayList
    end sub

    ' methods
    public sub LoadAll()
        dim oCon
        set oCon = new SiteSettingsDataConnection

        dim results
        results = oCon.GetAllSiteSettingsBySite(MY_SITE_GUID)

        if UBound(results) > 0 then
            for i = 0 to UBound(results, 2)
                dim mySetting
                set mySetting = new SiteSetting

                'mySetting.SiteGuid = results(SITESETTING_INDEX_SITEGUID, i)
                mySetting.Key = results(SITESETTING_INDEX_KEY, i)
                mySetting.Value = results(SITESETTING_INDEX_VALUE, i)
    
                mo_List.Add mySetting
            next
        end if
    end sub

    public function ToDictionary()
        dim myDict
        set myDict = Server.CreateObject("Scripting.Dictionary")

        for i = 0 to Count - 1
            dim mySetting
            set mySetting = mo_List(i)

            myDict.Add mySetting.Key, mySetting.Value
        next

        set ToDictionary = myDict
    end function

    ' properties
    public property get Count
        Count = mo_List.Count
    end property

    Public Default Property Get Item(index)
        set Item = mo_List(index)
    end property
end class
%>