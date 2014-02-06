<%
class Site
    ' attributes
    private ms_siteGuid, ms_name, ms_adminEmail, mb_active, _
            ms_fromEmailAddress, ms_fromName, mt_mondayOpen, _
            mt_mondayClose, mt_tuesdayOpen, mt_tuestdayClose, _
            mt_wednesdayOpen, mt_wednesdayClose, mt_thursdayOpen, _
            mt_thursdayClose, mt_fridayOpen, mt_fridayClose, _
            mt_saturdayOpen, mt_saturdayClose, mt_sundayOpen, _
            mt_sundayClose, ms_HomeUrl, ms_VantoraUrl, _
            ms_Address, ms_City, ms_State, _
            ms_ZipCode, ms_PhoneNumber, ms_DirectoryName,_
            ms_Country
    'ctor

    'methods
    public sub Load(sSiteGuid)
        if len(sSiteGuid) = 0 then exit sub

        dim myCon
        set myCon = new SiteDataConnection

        dim results
        results = myCon.GetSite(sSiteGuid)

        if UBound(results) > 0 then
            SiteGuid = results(SITE_INDEX_GUID, 0)
            Name = results(SITE_INDEX_NAME, 0)
            AdminEmail = results(SITE_INDEX_ADMINEMAIL, 0)
            'Active = results(SITE_INDEX_ACTIVE, 0)
            FromEmailAddress = results(SITE_INDEX_FROMEMAILADDRESS, 0)
            FromName = results(SITE_INDEX_FROMNAME, 0)
            MondayOpenTime = results(SITE_INDEX_MONDAYOPEN, 0)
            MondayCloseTime = results(SITE_INDEX_MONDAYCLOSE, 0)
            TuesdayOpenTime = results(SITE_INDEX_TUESDAYOPEN, 0)
            TuesdayCloseTime = results(SITE_INDEX_TUESDAYCLOSE, 0)
            WednesdayOpenTime = results(SITE_INDEX_WEDNESDAYOPEN, 0)
            WednesdayCloseTime = results(SITE_INDEX_WEDNESDAYCLOSE, 0)
            ThursdayOpenTime = results(SITE_INDEX_THURSDAYOPEN, 0)
            ThursdayCloseTime = results(SITE_INDEX_THURSDAYCLOSE, 0)
            FridayOpenTime = results(SITE_INDEX_FRIDAYOPEN, 0)
            FridayCloseTime = results(SITE_INDEX_FRIDAYCLOSE, 0)
            SaturdayOpenTime = results(SITE_INDEX_SATURDAYOPEN, 0)
            SaturdayCloseTime = results(SITE_INDEX_SATURDAYCLOSE, 0)
            SundayOpenTime = results(SITE_INDEX_SUNDAYOPEN, 0)
            SundayCloseTime = results(SITE_INDEX_SUNDAYCLOSE, 0)
            HomeUrl = results(SITE_INDEX_HOMEURL, 0)
            VantoraUrl = results(SITE_INDEX_VANTORAURL, 0)
            Address = results(SITE_INDEX_ADDRESS, 0)
            City = results(SITE_INDEX_CITY, 0)
            State = results(SITE_INDEX_STATE, 0)
            ZipCode = results(SITE_INDEX_ZIPCODE, 0)
            PhoneNumber = results(SITE_INDEX_PHONENUMBER, 0)
            DirectoryName = results(SITE_INDEX_DIRECTORYNAME, 0)
            Country = results(SITE_INDEX_COUNTRY, 0)
        end if
    end sub

    public sub LoadByName(sName)
        dim myCon
        set myCon = new SiteDataConnection

        dim results
        results = myCon.GetSiteByName(sName)

        if UBound(results) > 0 then
            SiteGuid = results(SITE_INDEX_GUID, 0)
            Name = results(SITE_INDEX_NAME, 0)
            AdminEmail = results(SITE_INDEX_ADMINEMAIL, 0)
            'Active = results(SITE_INDEX_ACTIVE, 0)
            FromEmailAddress = results(SITE_INDEX_FROMEMAILADDRESS, 0)
            FromName = results(SITE_INDEX_FROMNAME, 0)
            MondayOpenTime = results(SITE_INDEX_MONDAYOPEN, 0)
            MondayCloseTime = results(SITE_INDEX_MONDAYCLOSE, 0)
            TuesdayOpenTime = results(SITE_INDEX_TUESDAYOPEN, 0)
            TuesdayCloseTime = results(SITE_INDEX_TUESDAYCLOSE, 0)
            WednesdayOpenTime = results(SITE_INDEX_WEDNESDAYOPEN, 0)
            WednesdayCloseTime = results(SITE_INDEX_WEDNESDAYCLOSE, 0)
            ThursdayOpenTime = results(SITE_INDEX_THURSDAYOPEN, 0)
            ThursdayCloseTime = results(SITE_INDEX_THURSDAYCLOSE, 0)
            FridayOpenTime = results(SITE_INDEX_FRIDAYOPEN, 0)
            FridayCloseTime = results(SITE_INDEX_FRIDAYCLOSE, 0)
            SaturdayOpenTime = results(SITE_INDEX_SATURDAYOPEN, 0)
            SaturdayCloseTime = results(SITE_INDEX_SATURDAYCLOSE, 0)
            SundayOpenTime = results(SITE_INDEX_SUNDAYOPEN, 0)
            SundayCloseTime = results(SITE_INDEX_SUNDAYCLOSE, 0)
            HomeUrl = results(SITE_INDEX_HOMEURL, 0)
            VantoraUrl = results(SITE_INDEX_VANTORAURL, 0)
            Address = results(SITE_INDEX_ADDRESS, 0)
            City = results(SITE_INDEX_CITY, 0)
            State = results(SITE_INDEX_STATE, 0)
            ZipCode = results(SITE_INDEX_ZIPCODE, 0)
            PhoneNumber = results(SITE_INDEX_PHONENUMBER, 0)
            DirectoryName = results(SITE_INDEX_DIRECTORYNAME, 0)
            Country = results(SITE_INDEX_COUNTRY, 0)
        end if
    end sub

    public sub Save()
        dim myCon
        set myCon = new SiteDataConnection

        myCon.SaveSite SiteGuid, Name, AdminEmail, FromEmailAddress, FromName, MondayOpenTime, MondayCloseTime, TuesdayOpenTime, TuesdayCloseTime, _
                       WednesdayOpenTime, WednesdayCloseTime, ThursdayOpenTime, ThursdayCloseTime, FridayOpenTime, FridayCloseTime, _
                       SaturdayOpenTime, SaturdayCloseTime, SundayOpenTime, SundayCloseTime, HomeUrl, VantoraUrl, Address, City, State, ZipCode, _
                       PhoneNumber, DirectoryName, Country
    end sub

    'properties
    public property Get SiteGuid
        SiteGuid = ms_siteGuid
    end property

    public property Let SiteGuid(value)
        ms_siteGuid = value
    end property

    public property Get Name
        Name = ms_name
    end property

    public property Let Name(value)
        ms_name = value
    end property

    public property Get AdminEmail
        AdminEmail = ms_adminEmail
    end property

    public property Let AdminEmail(value)
        ms_adminEmail = value
    end property

    public property Get Active
        Active = mb_Active
    end property

    public property Get FromEmailAddress
        FromEmailAddress = ms_fromEmailAddress
    end property

    public property Let FromEmailAddress(value)
        ms_fromEmailAddress = value
    end property

    public property Get FromName
        FromName = ms_fromName
    end property

    public property Let FromName(value)
        ms_fromName = value
    end property

    public property Get MondayOpenTime
        MondayOpenTime = mt_mondayOpen
    end property

    public property Let MondayOpenTime(value)
        mt_mondayOpen = value
    end property

    public property Get MondayOpenTimeShortFormat
        MondayOpenTimeShortFormat = GetMilitaryTime(mt_mondayOpen)
    end property
    
    public property Get MondayCloseTime
        MondayCloseTime = mt_mondayClose
    end property

    public property Let MondayCloseTime(value)
        mt_mondayClose = value
    end property

    public property Get MondayCloseTimeShortFormat
        MondayCloseTimeShortFormat = GetMilitaryTime(mt_mondayClose)
    end property

    public property Get TuesdayOpenTime
        TuesdayOpenTime = mt_tuesdayOpen
    end property

    public property Let TuesdayOpenTime(value)
        mt_tuesdayOpen = value
    end property

    public property Get TuesdayOpenTimeShortFormat
        TuesdayOpenTimeShortFormat = GetMilitaryTime(mt_tuesdayOpen)
    end property

    public property Get TuesdayCloseTime
        TuesdayCloseTime = mt_tuestdayClose
    end property

    public property Let TuesdayCloseTime(value)
        mt_tuestdayClose = value
    end property

    public property Get TuesdayCloseTimeShortFormat
        TuesdayCloseTimeShortFormat = GetMilitaryTime(mt_tuestdayClose)
    end property

    public property Get WednesdayOpenTime
        WednesdayOpenTime = mt_wednesdayOpen
    end property

    public property Let WednesdayOpenTime(value)
        mt_wednesdayOpen = value
    end property

    public property Get WednesdayOpenTimeShortFormat
        WednesdayOpenTimeShortFormat = GetMilitaryTime(mt_wednesdayOpen)
    end property

    public property Get WednesdayCloseTime
        WednesdayCloseTime = mt_wednesdayClose
    end property

    public property Let WednesdayCloseTime(value)
        mt_wednesdayClose = value
    end property

    public property Get WednesdayCloseTimeShortFormat
        WednesdayCloseTimeShortFormat = GetMilitaryTime(mt_wednesdayClose)
    end property

    public property Get ThursdayOpenTime
        ThursdayOpenTime = mt_thursdayOpen
    end property

    public property Let ThursdayOpenTime(value)
        mt_thursdayOpen = value
    end property

    public property Get ThursdayOpenTimeShortFormat
        ThursdayOpenTimeShortFormat = GetMilitaryTime(mt_thursdayOpen)
    end property

    public property Get ThursdayCloseTime
        ThursdayCloseTime = mt_thursdayClose
    end property

    public property Let ThursdayCloseTime(value)
        mt_thursdayClose = value
    end property

    public property Get ThursdayCloseTimeShortFormat
        ThursdayCloseTimeShortFormat = GetMilitaryTime(mt_thursdayClose)
    end property

    public property Get FridayOpenTime
        FridayOpenTime = mt_fridayOpen
    end property

    public property Let FridayOpenTime(value)
        mt_fridayOpen = value
    end property

    public property Get FridayOpenTimeShortFormat
        FridayOpenTimeShortFormat = GetMilitaryTime(mt_fridayOpen)
    end property

    public property Get FridayCloseTime
        FridayCloseTime = mt_fridayClose
    end property

    public property Let FridayCloseTime(value)
        mt_fridayClose = value
    end property

    public property Get FridayCloseTimeShortFormat
        FridayCloseTimeShortFormat = GetMilitaryTime(mt_fridayClose)
    end property

    public property Get SaturdayOpenTime
        SaturdayOpenTime = mt_saturdayOpen
    end property

    public property Let SaturdayOpenTime(value)
        mt_saturdayOpen = value
    end property

    public property Get SaturdayOpenTimeShortFormat
        SaturdayOpenTimeShortFormat = GetMilitaryTime(mt_saturdayOpen)
    end property

    public property Get SaturdayCloseTime
        SaturdayCloseTime = mt_saturdayClose
    end property

    public property Let SaturdayCloseTime(value)
        mt_saturdayClose = value
    end property

    public property Get SaturdayCloseTimeShortFormat
        SaturdayCloseTimeShortFormat = GetMilitaryTime(mt_saturdayClose)
    end property

    public property Get SundayOpenTime
        SundayOpenTime = mt_sundayOpen
    end property

    public property Let SundayOpenTime(value)
        mt_sundayOpen = value
    end property

    public property Get SundayOpenTimeShortFormat
        SundayOpenTimeShortFormat = GetMilitaryTime(mt_sundayOpen)
    end property

    public property Get SundayCloseTime
        SundayCloseTime = mt_sundayClose
    end property

    public property Let SundayCloseTime(value)
        mt_sundayClose = value
    end property

    public property Get SundayCloseTimeShortFormat
        SundayCloseTimeShortFormat = GetMilitaryTime(mt_sundayClose)
    end property

    public property get HomeUrl
        HomeUrl = ms_HomeUrl
    end property

    public property let HomeUrl(value)
        ms_HomeUrl = value
    end property

    public property get VantoraUrl
        VantoraUrl = ms_VantoraUrl
    end property

    public property let VantoraUrl(value)
        ms_VantoraUrl = value
    end property

    public property Get Address
        Address = ms_Address
    end property

    public property let Address(value)
        ms_Address = value
    end property

    public property get City
        City = ms_City
    end property

    public property let City(value)
        ms_City = value
    end property

    public property get State
        State = ms_State
    end property

    public property let State(value)
        ms_State = value
    end property

    public property get ZipCode
        ZipCode = ms_ZipCode
    end property

    public property let ZipCode(value)
        ms_ZipCode = value
    end property

    public property Get PhoneNumber
        PhoneNumber = ms_PhoneNumber
    end property

    public property let PhoneNumber(value)
        ms_PhoneNumber = value
    end property

    public property get DirectoryName
        DirectoryName = ms_DirectoryName
    end property

    public property let DirectoryName(value)
        ms_DirectoryName = value
    end property

    public property get Country
        Country = ms_Country
    end property

    public property let Country(value)
        ms_Country = value
    end property
end class
%>