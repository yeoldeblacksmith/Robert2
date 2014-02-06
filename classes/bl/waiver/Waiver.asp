 <%
class Waiver
    ' attributes
    dim mn_WaiverId, mn_PlayerId, mn_EventId, _
        mdt_CreateDateTime, mn_ParentId, ms_PhoneNumber, _
        ms_Address, ms_City, ms_State, _
        ms_ZipCode, ms_EmailAddress, mn_LegaleseId, _
        ms_Signature, ms_HashId, mb_Loaded, mb_Exported, ms_IpAddress

    ' ctor
    public sub Class_Initialize
        mb_Loaded = false
    end sub

    ' methods

    public sub Delete(sHashId)
        dim myCon
        set myCon = new WaiverDataConnection

        myCon.DeleteWaiver sHashId
    end sub

    public sub LoadById(nWaiverId)
        dim myCon
        set myCon = new WaiverDataConnection

        dim results
        results = myCon.GetWaiversById(nWaiverId)

        if UBound(results) > 0 then
            WaiverId = results(W2_COLUMN_INDEX_WAIVERID, 0)
            SiteGuid = results(W2_COLUMN_INDEX_SITEGUID, 0)
            CreateDateTime = results(W2_COLUMN_INDEX_CREATEDATETIME, 0)
            PlayerId = results(W2_COLUMN_INDEX_PLAYERID, 0)
            ParentId = results(W2_COLUMN_INDEX_PARENTID, 0)
            PhoneNumber = results(W2_COLUMN_INDEX_PHONENUMBER, 0)
            Address = results(W2_COLUMN_INDEX_ADDRESS, 0)
            City = results(W2_COLUMN_INDEX_CITY, 0)
            State = results(W2_COLUMN_INDEX_STATE, 0)
            ZipCode = results(W2_COLUMN_INDEX_ZIPCODE, 0)
            EmailAddress = results(W2_COLUMN_INDEX_EMAILADDRESS, 0)
            LegaleseId = results(W2_COLUMN_INDEX_EMAILADDRESS, 0)
            HashId = results(W2_COLUMN_INDEX_HASHID, 0)
            SubmittedFromIpAddress = results(W2_COLUMN_INDEX_SUBMITTEDFROMIPADDRESS, 0)
            Signature = results(W2_COLUMN_INDEX_SIGNATURE, 0)
            Exported = results(W2_COLUMN_INDEX_EXPORTED, 0)

            mb_Loaded = true
        end if

    end sub

    public sub Load(sHashId)
        dim myCon
        set myCon = new WaiverDataConnection

        dim results
        results = myCon.GetWaiversByHashId(sHashId)

        if UBound(results) > 0 then
            WaiverId = results(W2_COLUMN_INDEX_WAIVERID, 0)
            SiteGuid = results(W2_COLUMN_INDEX_SITEGUID, 0)
            CreateDateTime = results(W2_COLUMN_INDEX_CREATEDATETIME, 0)
            PlayerId = results(W2_COLUMN_INDEX_PLAYERID, 0)
            ParentId = results(W2_COLUMN_INDEX_PARENTID, 0)
            PhoneNumber = results(W2_COLUMN_INDEX_PHONENUMBER, 0)
            Address = results(W2_COLUMN_INDEX_ADDRESS, 0)
            City = results(W2_COLUMN_INDEX_CITY, 0)
            State = results(W2_COLUMN_INDEX_STATE, 0)
            ZipCode = results(W2_COLUMN_INDEX_ZIPCODE, 0)
            EmailAddress = results(W2_COLUMN_INDEX_EMAILADDRESS, 0)
            LegaleseId = results(W2_COLUMN_INDEX_EMAILADDRESS, 0)
            HashId = results(W2_COLUMN_INDEX_HASHID, 0)
            SubmittedFromIpAddress = results(W2_COLUMN_INDEX_SUBMITTEDFROMIPADDRESS, 0)
            Signature = results(W2_COLUMN_INDEX_SIGNATURE, 0)
            Exported = results(W2_COLUMN_INDEX_EXPORTED, 0)

            mb_Loaded = true
        end if

    end sub

    public sub Save()
        dim myCon, md5
        set myCon = new WaiverDataConnection
        set md5 = new Md5Generator

        if len(HashId) = 0 then HashId = md5.CreateHash(FirstName & WaiverId & LastName & FormatDateTime(Now()) & PlayerDOB & "52029")

        WaiverId = myCon.SaveWaiver(WaiverId, PlayerId, ParentId, _
                                    PhoneNumber, Address, City, State, _
                                    ZipCode, CreateDateTime, EmailAddress, _
                                    Signature, HashId, SubmittedFromIpAddress)
    end sub



    ' properties

    public property get Address
        Address = ms_Address
    end property

    public property let Address(value)
        ms_Address = value
    end property

    public property get AgeAtSigning
        dim n_AgeAtSigning
        If DatePart("y",PlayerDOB) <= DatePart("y",mdt_CreateDateTime) Then 
            n_AgeAtSigning = DateDiff("yyyy", PlayerDOB, mdt_CreateDateTime) 
        Else 
            n_AgeAtSigning = DateDiff("yyyy",PlayerDOB,DateAdd("yyyy",-1,mdt_CreateDateTime)) 
        End If
        AgeAtSigning = n_AgeAtSigning
    end property

    public property get City
        City = ms_City
    end property

    public property let City(value)
        ms_City = value
    end property

    public property get CheckedInToday
        dim myPlayDate
        set myPlayDate = new PlayDateTime
        myPlayDate.GetTodaysCheckInStatus PlayerId
        
        If IsNull(myPlayDate.CheckInDateTimeAdmin) OR IsEmpty(myPlayDate.CheckInDateTimeAdmin) OR Len(myPlayDate.CheckInDateTimeAdmin) = 0 Then
            CheckedInToday = false
        Else
            CheckInToday = true
        End If
    end property

    public property get CreateDateTime
        CreateDateTime = mdt_CreateDateTime
    end property

    public property let CreateDateTime(value)
        mdt_CreateDateTime = value
    end property

    public property get CurrentAge
        Dim n_Age

        If DatePart("y",PlayerDOB) <= DatePart("y",now()) Then 
            n_Age = DateDiff("yyyy", PlayerDOB, now()) 
        Else 
            n_Age = DateDiff("yyyy",PlayerDOB,DateAdd("yyyy",-1,now())) 
        End If
        CurrentAge = n_Age
    end property 

    public property get EmailAddress
        EmailAddress = ms_EmailAddress
    end property

    public property let EmailAddress(value)
        ms_EmailAddress = value
    end property

    public property get EmailList
        dim myPlayer
        set myPlayer = New Player
        myPlayer.LoadById PlayerId
        EmailList = myPlayer.EmailList
    end property

    public property get Exported
        Exported = mb_Exported
    end property

    public property let Exported(value)
        mb_Exported = value
    end property

    public property get FirstName
        dim myPlayer
        set myPlayer = New Player
        myPlayer.LoadById PlayerId
        FirstName = myPlayer.FirstName
    end property

    public property get HashId
        HashId = ms_HashId
    end property

    public property let HashId(value)
        ms_HashId = value
    end property

    public property get LastName
        dim myPlayer
        set myPlayer = New Player
        myPlayer.LoadById PlayerId
        LastName = myPlayer.LastName
    end property

    public property get LegaleseId
        LegaleseId = mn_LegaleseId
    end property

    public property let LegaleseId(value)
        mn_LegaleseId = value
    end property

    public property get Loaded
        Loaded = mb_Loaded
    end property

    public property get ParentId
        ParentId = mn_ParentId
    end property

    public property let ParentId(value)
        mn_ParentId = value
    end property

    public property get PhoneNumber
        PhoneNumber = ms_PhoneNumber
    end property

    public property let PhoneNumber(value)
        ms_PhoneNumber = value
    end property

    public property get PlayerDOB
        dim myPlayer
        set myPlayer = New Player
        myPlayer.LoadById PlayerId
        PlayerDOB = myPlayer.DateOfBirth
    end property

    public property get ParentDOB
        dim myPlayer
        set myPlayer = New Player
        myPlayer.LoadById ParentId
        ParentDOB = myPlayer.DateOfBirth
    end property

    public property get ParentFirstName
        dim myPlayer
        set myPlayer = New Player
        myPlayer.LoadById ParentId
        ParentFirstName = myPlayer.FirstName
    end property

    public property get ParentLastName
        dim myPlayer
        set myPlayer = New Player
        myPlayer.LoadById ParentId
        ParentLastName = myPlayer.LastName
    end property

    public property get PlayerId
        PlayerId = mn_PlayerId
    end property

    public property let PlayerId(value)
        mn_PlayerId = value
    end property

    public property get Signature
        Signature = ms_Signature
    end property

    public property let Signature(value)
        ms_Signature = value
    end property

    public property get State
        State = ms_State
    end property

    public property let State(value)
        ms_State = value
    end property

    public property get SubmittedFromIpAddress
        SubmittedFromIpAddress = ms_IpAddress
    end property

    public property let SubmittedFromIpAddress(value)
        ms_IpAddress = value
    end property

    public property get ValidityStatus
        dim n_AgeAtSigning
        dim n_ValidWaiver: n_ValidWaiver = WAIVER_VALIDITYSTATUS_GOOD

        If NOT (Year(mdt_CreateDateTime) = Year(Now())) then
            n_ValidWaiver = WAIVER_VALIDITYSTATUS_EXPIRED
        end if

        If CurrentAge >= 18 Then
            If DatePart("y",PlayerDOB) <= DatePart("y",mdt_CreateDateTime) Then 
                n_AgeAtSigning = DateDiff("yyyy", PlayerDOB, mdt_CreateDateTime) 
            Else 
                n_AgeAtSigning = DateDiff("yyyy",PlayerDOB,DateAdd("yyyy",-1,mdt_CreateDateTime)) 
            End If
            
            If n_AgeAtSigning <= 17 Then
                n_ValidWaiver = WAIVER_VALIDITYSTATUS_NEWADULT
            End If            
        End If 

        ValidityStatus = n_validwaiver
    end property

    public property get IsValid
        Dim bIsValid
        If ValidityStatus = WAIVER_VALIDITYSTATUS_GOOD Then
            bIsValid = True
        Else
            bIsValid = False
        End If

        IsValid = bIsValid
    end property

    public property get WaiverId
        WaiverId = mn_WaiverId
    end property

    public property let WaiverId(value)
        mn_WaiverId = value
    end property

    public property get ZipCode
        ZipCode = ms_ZipCode
    end property

    public property let ZipCode(value)
        ms_ZipCode = value
    end property




end class
%>