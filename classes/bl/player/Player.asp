 <%
class Player
    ' attributes
    dim mn_PlayerId, ms_SiteGuid, ms_FirstName, ms_LastName, mb_Loaded,  _
        mdt_DateOfBirth, ms_PhoneNumber, ms_Address, _
        ms_City, ms_State, ms_ZipCode, ms_EmailAddress, mb_EmailList

    ' ctor
    public sub Class_Initialize
        mb_Loaded = false
    end sub

    ' methods

    public sub LoadById(nPlayerId)
        dim myCon
        set myCon = new PlayerDataConnection

        dim results
        results = myCon.GetPlayerById(nPlayerId)

        if UBound(results) > 0 then
            SiteGuid = results(PLAYER_INDEX_SITEGUID, 0)
            PlayerId = results(PLAYER_INDEX_PLAYERID, 0)
            FirstName = results(PLAYER_INDEX_FIRSTNAME, 0)
            LastName = results(PLAYER_INDEX_LASTNAME, 0)
            DateOfBirth = results(PLAYER_INDEX_DATEOFBIRTH, 0)
            PhoneNumber = results(PLAYER_INDEX_PHONENUMBER, 0)
            Address = results(PLAYER_INDEX_ADDRESS, 0)
            City = results(PLAYER_INDEX_CITY, 0)
            State = results(PLAYER_INDEX_STATE, 0)
            ZipCode = results(PLAYER_INDEX_ZIPCODE, 0)
            EmailAddress = results(PLAYER_INDEX_EMAILADDRESS, 0)
            EmailList = results(PLAYER_INDEX_EMAILLIST, 0)

            mb_Loaded = true
        end if
    end sub

    public sub LoadByLongInfo(sFirstname, sLastname, dDOB, sAddress)
        dim myCon
        set myCon = new PlayerDataConnection

        dim results
        results = myCon.GetPlayerByLongInfo(sFirstname, sLastname, dDOB, sAddress)

        if UBound(results) > 0 then
            SiteGuid = results(PLAYER_INDEX_SITEGUID, 0)
            PlayerId = results(PLAYER_INDEX_PLAYERID, 0)
            FirstName = results(PLAYER_INDEX_FIRSTNAME, 0)
            LastName = results(PLAYER_INDEX_LASTNAME, 0)
            DateOfBirth = results(PLAYER_INDEX_DATEOFBIRTH, 0)
            PhoneNumber = results(PLAYER_INDEX_PHONENUMBER, 0)
            Address = results(PLAYER_INDEX_ADDRESS, 0)
            City = results(PLAYER_INDEX_CITY, 0)
            State = results(PLAYER_INDEX_STATE, 0)
            ZipCode = results(PLAYER_INDEX_ZIPCODE, 0)
            EmailAddress = results(PLAYER_INDEX_EMAILADDRESS, 0)
            EmailList = results(PLAYER_INDEX_EMAILLIST, 0)

            mb_Loaded = true
        end if
    end sub

    public sub Save()
        dim myCon
        set myCon = new PlayerDataConnection

        PlayerId = myCon.SavePlayer(PlayerId, EmailList, FirstName, LastName, _
                                    DateOfBirth, PhoneNumber, _
                                    Address, City, State, _
                                    ZipCode, EmailAddress)
    end sub

    ' properties

    public property get SiteGuid
        SiteGuid = ms_SiteGuid
    end property

    public property let SiteGuid(value)
        ms_SiteGuid = value
    end property

    public property get PlayerId
        PlayerId = mn_PlayerId
    end property

    public property let PlayerId(value)
        mn_PlayerId = value
    end property

    public property get FirstName
        FirstName = ms_FirstName
    end property

    public property let FirstName(value)
        ms_FirstName = value
    end property

    public property get LastName
        LastName = ms_LastName
    end property

    public property let LastName(value)
        ms_LastName = value
    end property

    public property get DateOfBirth
        DateOfBirth = mdt_DateOfBirth
    end property

    public property let DateOfBirth(value)
        mdt_DateOfBirth = value
    end property

    public property get PhoneNumber
        PhoneNumber = ms_PhoneNumber
    end property

    public property let PhoneNumber(value)
        ms_PhoneNumber = value
    end property

    public property get Address
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

    public property get EmailAddress
        EmailAddress = ms_EmailAddress
    end property

    public property let EmailAddress(value)
        ms_EmailAddress = value
    end property

    public property get Loaded
        Loaded = mb_Loaded
    end property

    public property get EmailList
        EmailList = mb_EmailList
    end property

    public property let EmailList(value)
        mb_EmailList = value
    end property


end class
%>