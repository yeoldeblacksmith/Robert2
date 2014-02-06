<%
class PlayerCollection
    private mo_List

    'ctor
    public sub Class_Initialize
        set mo_List = new ArrayList
    end sub

    ' methods
    public sub Add(value)
        mo_List.Add value
    end sub

    public sub LoadPlayersByNameDOB(sFirstName, sLastName, dtDOB)
        dim myCon
        set myCon = new PlayerDataConnection

        dim playersColl
        playersColl = myCon.GetPlayersByNameDOB(sFirstName, sLastName, dtDOB) 

        if UBound(playersColl) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(playersColl, 2) 
                dim myPlayer
                set myPlayer = new Player

                myPlayer.PlayerId = playersColl(PLAYER_INDEX_PLAYERID, i)
                myPlayer.SiteGuid = playersColl(PLAYER_INDEX_SITEGUID, i)
                myPlayer.FirstName = playersColl(PLAYER_INDEX_FIRSTNAME, i)
                myPlayer.LastName = playersColl(PLAYER_INDEX_LASTNAME, i)
                myPlayer.DateOfBirth = playersColl(PLAYER_INDEX_DATEOFBIRTH, i)
                myPlayer.PhoneNumber = playersColl(PLAYER_INDEX_PHONENUMBER, i)
                myPlayer.Address = playersColl(PLAYER_INDEX_ADDRESS, i)
                myPlayer.City = playersColl(PLAYER_INDEX_CITY, i)
                myPlayer.State = playersColl(PLAYER_INDEX_STATE, i)
                myPlayer.ZipCode = playersColl(PLAYER_INDEX_ZIPCODE, i)
                myPlayer.EmailAddress = playersColl(PLAYER_INDEX_EMAILADDRESS, i)
                myPlayer.EmailList = playersColl(PLAYER_INDEX_EMAILLIST, i)             

                mo_List.Add myPlayer
            next
        end if

    end sub

    public sub LoadPlayersByNameDOBAddress(sFirstName, sLastName, dtDOB, sAddress)
        dim myCon
        set myCon = new PlayerDataConnection

        dim playersColl
        playersColl = myCon.GetPlayersByNameDOBAddress(sFirstName, sLastName, dtDOB, sAddress) 

        if UBound(playersColl) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(playersColl, 2) 
                dim myPlayer
                set myPlayer = new Player

                myPlayer.PlayerId = playersColl(PLAYER_INDEX_PLAYERID, i)
                myPlayer.SiteGuid = playersColl(PLAYER_INDEX_SITEGUID, i)
                myPlayer.FirstName = playersColl(PLAYER_INDEX_FIRSTNAME, i)
                myPlayer.LastName = playersColl(PLAYER_INDEX_LASTNAME, i)
                myPlayer.DateOfBirth = playersColl(PLAYER_INDEX_DATEOFBIRTH, i)
                myPlayer.PhoneNumber = playersColl(PLAYER_INDEX_PHONENUMBER, i)
                myPlayer.Address = playersColl(PLAYER_INDEX_ADDRESS, i)
                myPlayer.City = playersColl(PLAYER_INDEX_CITY, i)
                myPlayer.State = playersColl(PLAYER_INDEX_STATE, i)
                myPlayer.ZipCode = playersColl(PLAYER_INDEX_ZIPCODE, i)
                myPlayer.EmailAddress = playersColl(PLAYER_INDEX_EMAILADDRESS, i)
                myPlayer.EmailList = playersColl(PLAYER_INDEX_EMAILLIST, i)             

                mo_List.Add myPlayer
            next
        end if

    end sub


    public sub LoadPlayersStreetChooser()
        dim myCon
        set myCon = new PlayerDataConnection

        dim playersColl
        playersColl = myCon.GetRandomStreets() 

        if UBound(playersColl) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(playersColl, 2) 
                dim myPlayer
                set myPlayer = new Player

                myPlayer.PlayerId = playersColl(PLAYER_INDEX_PLAYERID, i)
                myPlayer.SiteGuid = playersColl(PLAYER_INDEX_SITEGUID, i)
                myPlayer.FirstName = playersColl(PLAYER_INDEX_FIRSTNAME, i)
                myPlayer.LastName = playersColl(PLAYER_INDEX_LASTNAME, i)
                myPlayer.DateOfBirth = playersColl(PLAYER_INDEX_DATEOFBIRTH, i)
                myPlayer.PhoneNumber = playersColl(PLAYER_INDEX_PHONENUMBER, i)
                myPlayer.Address = playersColl(PLAYER_INDEX_ADDRESS, i)
                myPlayer.City = playersColl(PLAYER_INDEX_CITY, i)
                myPlayer.State = playersColl(PLAYER_INDEX_STATE, i)
                myPlayer.ZipCode = playersColl(PLAYER_INDEX_ZIPCODE, i)
                myPlayer.EmailAddress = playersColl(PLAYER_INDEX_EMAILADDRESS, i)
                myPlayer.EmailList = playersColl(PLAYER_INDEX_EMAILLIST, i)             

                mo_List.Add myPlayer
            next
        end if

    end sub



  public sub Remove(value)
        mo_List.Remove value
    end sub

    'properties
    public property get Count
        Count = mo_List.Count
    end property

    Public Default Property Get Item(index)
        set Item = mo_List(index)
    end property
end class
%>