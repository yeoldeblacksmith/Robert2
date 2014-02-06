 <%
class WaiverForExport
    ' attributes
    dim mn_WaiverId, ms_FirstName, ms_LastName, mn_Age, mdt_DateOfBirth, ms_PhoneNumber, _
        ms_Address, ms_City, ms_State, ms_ZipCode, mdt_PlayDate, mdt_CreateDateTime, _
        ms_EmailAddress,  mb_EmailList, mn_EventId, ms_HashId, ms_ParentFirstName, _
        ms_ParentLastName, mdt_PlayTime, ms_IpAddress, ms_WaiverCustomFields, mb_Loaded, mb_Exported

    ' ctor
    public sub Class_Initialize
        mb_Loaded = false
    end sub

    ' methods
    public function ToString()
        dim str: str = ""

        if isnull(FirstName) then FirstName = ""
        if isnull(LastName) then LastName = ""
        if isnull(ParentFirstName) then ParentFirstName = ""
        if isnull(ParentLastName) then ParentLastName = ""
        if isnull(Address) then Address = ""
        if isnull(City) then City = ""
        if isnull(ZipCode) then ZipCode = ""
        if isnull(EmailAddress) then EmailAddress = ""

        str = """" & WaiverId & """," & _
              """" & Replace(FirstName, chr(34), "'") & """," & _
              """" & Replace(LastName ,chr(34), "'") & """," & _
              """" & Age & """," & _
              """" & DateOfBirth & """," & _
              """" & Replace(ParentFirstName, """", "'") & """," & _
              """" & Replace(ParentLastName, """", "'") & """," & _
              """" & PhoneNumber & """," & _
              """" & Replace(Address ,chr(34), "'") & """," & _
              """" & Replace(City ,chr(34), "'") & """," & _
              """" & State & """," & _
              """" & Replace(ZipCode ,chr(34), "'") & """," 

        for each column in arrColumnNames
              str = str & """" & fillCustomFieldsColumns(column, waivercustomfields) & ""","  
        next

        str = str & _
              """" & PlayDate & """," & _
              """" & GetTimeString(PlayTime) & """," & _
              """" & CreateDateTime & """," & _
              """" & Replace(EmailAddress ,chr(34), "'") & """," & _
              """" & EmailList & """"

        ToString = str
    end function

    ' properties

    public property get Address
        Address = ms_Address
    end property

    public property let Address(value)
        ms_Address = value
    end property

    public property get Age
        Age = mn_Age
    end property

    public property let Age(value)
        mn_Age = value
    end property

    public property get City
        City = ms_City
    end property

    public property let City(value)
        ms_City = value
    end property

    public property get CreateDateTime
        CreateDateTime = mdt_CreateDateTime
    end property

    public property let CreateDateTime(value)
        mdt_CreateDateTime = value
    end property

    public property get EmailAddress
        EmailAddress = ms_EmailAddress
    end property

    public property let EmailAddress(value)
        ms_EmailAddress = value
    end property

    public property get EmailList
        EmailList = mb_EmailList
    end property

    public property let EmailList(value)
        mb_EmailList = value
    end property

    public property get EventId
        EventId = mn_EventId
    end property

    public property let EventId(value)
        mn_EventId = value
    end property

    public property get Exported
        Exported = mb_Exported
    end property

    public property let Exported(value)
        mb_Exported = value
    end property

    public property get FirstName
        FirstName = ms_FirstName
    end property

    public property let FirstName(value)
        ms_FirstName = value
    end property

    public property get HashId
        HashId = ms_HashId
    end property

    public property let HashId(value)
        ms_HashId = value
    end property

    public property get LastName
        LastName = ms_LastName
    end property

    public property let LastName(value)
        ms_LastName = value
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

    public property get PlayDate
        PlayDate = mdt_PlayDate
    end property

    public property let PlayDate(value)
        mdt_PlayDate = value
    end property

    public property get PlayTime
        PlayTime = mdt_PlayTime
    end property

    public property let PlayTime(value)
        mdt_PlayTime = value
    end property

    public property get PhoneNumber
        PhoneNumber = ms_PhoneNumber
    end property

    public property let PhoneNumber(value)
        ms_PhoneNumber = value
    end property

    public property get DateOfBirth
        DateOfBirth = mdt_DateOfBirth
    end property

    public property let DateOfBirth(value)
        mdt_DateOfBirth = value
    end property

    public property get ParentFirstName
        ParentFirstName = ms_ParentFirstName
    end property

    public property let ParentFirstName(value)
        ms_ParentFirstName = value
    end property

    public property get ParentLastName
        ParentLastName = ms_ParentLastName
    end property

    public property let ParentLastName(value)
        ms_ParentLastName = value
    end property

    public property get PlayerId
        PlayerId = mn_PlayerId
    end property

    public property let PlayerId(value)
        mn_PlayerId = value
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

    public property get WaiverId
        WaiverId = mn_WaiverId
    end property

    public property let WaiverId(value)
        mn_WaiverId = value
    end property

    public property get WaiverCustomFields
        WaiverCustomFields = ms_WaiverCustomFields
    end property

    public property let WaiverCustomFields(value)
        ms_WaiverCustomFields = value
    end property

    public property get ZipCode
        ZipCode = ms_ZipCode
    end property

    public property let ZipCode(value)
        ms_ZipCode = value
    end property




end class
%>