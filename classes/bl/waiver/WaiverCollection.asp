<%
class WaiverCollection
    private mo_List, mn_ExportOption, mdt_ExportStart, _
            mdt_ExportEnd

        private EXPORTOPTION_NOTSET
        private EXPORTOPTION_ALL 
        private EXPORTOPTION_CUSTOM 
        private EXPORTOPTION_MONTH 
        private EXPORTOPTION_YEAR 

    'ctor
    public sub Class_Initialize
        set mo_List = new ArrayList
        
        EXPORTOPTION_NOTSET = -1
        EXPORTOPTION_ALL = 0
        EXPORTOPTION_CUSTOM = 1
        EXPORTOPTION_MONTH = 2
        EXPORTOPTION_YEAR = 3

        mn_ExportOption = EXPORTOPTION_NOTSET
    end sub

    ' methods
    public sub Add(value)
        mo_List.Add value
    end sub
    
    public sub LoadByEventId(nEventId, IncludeInvalid, IncludeCheckedIn)
        dim myCon
        set myCon = new WaiverDataConnection

        dim waivers
        waivers = myCon.GetWaiversByEventId(nEventId, IncludeInvalid, IncludeCheckedIn)

        if UBound(waivers) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(waivers, 2) 
                dim myWaiver
                set myWaiver = new Waiver

                myWaiver.WaiverId = waivers(W2_COLUMN_INDEX_WAIVERID, i)
                myWaiver.CreateDateTime = waivers(W2_COLUMN_INDEX_CREATEDATETIME, i)
                myWaiver.PlayerId = waivers(W2_COLUMN_INDEX_PLAYERID, i)
                myWaiver.ParentId = waivers(W2_COLUMN_INDEX_PARENTID, i)
                myWaiver.PhoneNumber = waivers(W2_COLUMN_INDEX_PHONENUMBER, i)
                myWaiver.Address = waivers(W2_COLUMN_INDEX_ADDRESS, i)
                myWaiver.City = waivers(W2_COLUMN_INDEX_CITY, i)
                myWaiver.State = waivers(W2_COLUMN_INDEX_STATE, i)
                myWaiver.ZipCode = waivers(W2_COLUMN_INDEX_ZIPCODE, i)
                myWaiver.EmailAddress = waivers(W2_COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.LegaleseId = waivers(W2_COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.HashId = waivers(W2_COLUMN_INDEX_HASHID, i)
                myWaiver.SubmittedFromIpAddress = waivers(W2_COLUMN_INDEX_SUBMITTEDFROMIPADDRESS, i)
                myWaiver.Signature = waivers(W2_COLUMN_INDEX_SIGNATURE, i)
                myWaiver.Exported = waivers(W2_COLUMN_INDEX_EXPORTED, i) 

                mo_List.Add myWaiver
            next
        end if
    end sub

    public sub LoadForExportByAll(IncludePrevious)
        dim myCon
        set myCon = new WaiverDataConnection

        dim waivers
        waivers = myCon.GetWaiversForExportByAll(IncludePrevious)

        if UBound(waivers) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(waivers, 2) 
                dim myWaiver       
                set myWaiver = new WaiverForExport

                myWaiver.WaiverId = waivers(COLUMN_INDEX_WAIVERID, i)
                myWaiver.FirstName = waivers(COLUMN_INDEX_FIRSTNAME, i)
                myWaiver.LastName = waivers(COLUMN_INDEX_LASTNAME, i)
                myWaiver.Age = waivers(COLUMN_INDEX_AGE, i)
                myWaiver.DateOfBirth = waivers(COLUMN_INDEX_DATEOFBIRTH, i)
                myWaiver.PhoneNumber = waivers(COLUMN_INDEX_PHONENUMBER, i)
                myWaiver.Address = waivers(COLUMN_INDEX_ADDRESS, i)
                myWaiver.City = waivers(COLUMN_INDEX_CITY, i)
                myWaiver.State = waivers(COLUMN_INDEX_STATE, i)
                myWaiver.ZipCode = waivers(COLUMN_INDEX_ZIPCODE, i)
                myWaiver.PlayDate = waivers(COLUMN_INDEX_PLAYDATE, i)
                myWaiver.CreateDateTime = waivers(COLUMN_INDEX_WAIVERDATE, i)
                myWaiver.EmailAddress = waivers(COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.EmailList = waivers(COLUMN_INDEX_EMAILLIST, i)
                myWaiver.EventId = waivers(COLUMN_INDEX_EVENTID, i)
                myWaiver.HashId = waivers(COLUMN_INDEX_HASHID, i)
                myWaiver.ParentFirstName = waivers(COLUMN_INDEX_PARENTFIRSTNAME, i)
                myWaiver.ParentLastName = waivers(COLUMN_INDEX_PARENTLASTNAME, i)
                myWaiver.PlayTime = waivers(COLUMN_INDEX_PLAYTIME, i)
                myWaiver.SubmittedFromIpAddress = waivers(COLUMN_INDEX_SUBMITTEDFROMIPADDRESS, i)
                myWaiver.WaiverCustomFields = waivers(COLUMN_INDEX_CUSTOMFIELDS, i)

                mo_List.Add myWaiver
            next
        end if

        mn_ExportOption = EXPORTOPTION_ALL
    end sub

    public sub LoadForExportByCustomRange(StartDate, EndDate, IncludePrevious)
        dim myCon
        set myCon = new WaiverDataConnection

        dim waivers
        waivers = myCon.GetWaiversForExportByCustomRange(StartDate, EndDate, IncludePrevious)

        if UBound(waivers) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(waivers, 2) 
                dim myWaiver
                set myWaiver = new WaiverForExport

                myWaiver.WaiverId = waivers(COLUMN_INDEX_WAIVERID, i)
                myWaiver.FirstName = waivers(COLUMN_INDEX_FIRSTNAME, i)
                myWaiver.LastName = waivers(COLUMN_INDEX_LASTNAME, i)
                myWaiver.Age = waivers(COLUMN_INDEX_AGE, i)
                myWaiver.DateOfBirth = waivers(COLUMN_INDEX_DATEOFBIRTH, i)
                myWaiver.PhoneNumber = waivers(COLUMN_INDEX_PHONENUMBER, i)
                myWaiver.Address = waivers(COLUMN_INDEX_ADDRESS, i)
                myWaiver.City = waivers(COLUMN_INDEX_CITY, i)
                myWaiver.State = waivers(COLUMN_INDEX_STATE, i)
                myWaiver.ZipCode = waivers(COLUMN_INDEX_ZIPCODE, i)
                myWaiver.PlayDate = waivers(COLUMN_INDEX_PLAYDATE, i)
                myWaiver.CreateDateTime = waivers(COLUMN_INDEX_WAIVERDATE, i)
                myWaiver.EmailAddress = waivers(COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.EmailList = waivers(COLUMN_INDEX_EMAILLIST, i)
                myWaiver.EventId = waivers(COLUMN_INDEX_EVENTID, i)
                myWaiver.HashId = waivers(COLUMN_INDEX_HASHID, i)
                myWaiver.ParentFirstName = waivers(COLUMN_INDEX_PARENTFIRSTNAME, i)
                myWaiver.ParentLastName = waivers(COLUMN_INDEX_PARENTLASTNAME, i)
                myWaiver.PlayTime = waivers(COLUMN_INDEX_PLAYTIME, i)
                myWaiver.SubmittedFromIpAddress = waivers(COLUMN_INDEX_SUBMITTEDFROMIPADDRESS, i)
                myWaiver.WaiverCustomFields = waivers(COLUMN_INDEX_CUSTOMFIELDS, i)


                mo_List.Add myWaiver
            next
        end if

        mn_ExportOption = EXPORTOPTION_CUSTOM
        mdt_ExportStart = StartDate
        mdt_ExportEnd = EndDate
    end sub

    public sub LoadForExportByMonth(IncludePrevious)
        dim myCon
        set myCon = new WaiverDataConnection

        dim waivers
        waivers = myCon.GetWaiversForExportByMonth(IncludePrevious)

        if UBound(waivers) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(waivers, 2) 
                dim myWaiver
                set myWaiver = new WaiverForExport

                myWaiver.WaiverId = waivers(COLUMN_INDEX_WAIVERID, i)
                myWaiver.FirstName = waivers(COLUMN_INDEX_FIRSTNAME, i)
                myWaiver.LastName = waivers(COLUMN_INDEX_LASTNAME, i)
                myWaiver.Age = waivers(COLUMN_INDEX_AGE, i)
                myWaiver.DateOfBirth = waivers(COLUMN_INDEX_DATEOFBIRTH, i)
                myWaiver.PhoneNumber = waivers(COLUMN_INDEX_PHONENUMBER, i)
                myWaiver.Address = waivers(COLUMN_INDEX_ADDRESS, i)
                myWaiver.City = waivers(COLUMN_INDEX_CITY, i)
                myWaiver.State = waivers(COLUMN_INDEX_STATE, i)
                myWaiver.ZipCode = waivers(COLUMN_INDEX_ZIPCODE, i)
                myWaiver.PlayDate = waivers(COLUMN_INDEX_PLAYDATE, i)
                myWaiver.CreateDateTime = waivers(COLUMN_INDEX_WAIVERDATE, i)
                myWaiver.EmailAddress = waivers(COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.EmailList = waivers(COLUMN_INDEX_EMAILLIST, i)
                myWaiver.EventId = waivers(COLUMN_INDEX_EVENTID, i)
                myWaiver.HashId = waivers(COLUMN_INDEX_HASHID, i)
                myWaiver.ParentFirstName = waivers(COLUMN_INDEX_PARENTFIRSTNAME, i)
                myWaiver.ParentLastName = waivers(COLUMN_INDEX_PARENTLASTNAME, i)
                myWaiver.PlayTime = waivers(COLUMN_INDEX_PLAYTIME, i)
                myWaiver.SubmittedFromIpAddress = waivers(COLUMN_INDEX_SUBMITTEDFROMIPADDRESS, i)
                myWaiver.WaiverCustomFields = waivers(COLUMN_INDEX_CUSTOMFIELDS, i)

                mo_List.Add myWaiver
            next
        end if

        mn_ExportOption = EXPORTOPTION_MONTH
    end sub

    public sub LoadForExportByYear(IncludePrevious)
        dim myCon
        set myCon = new WaiverDataConnection

        dim waivers
        waivers = myCon.GetWaiversForExportByYear(IncludePrevious)

        if UBound(waivers) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(waivers, 2) 
                dim myWaiver
                set myWaiver = new WaiverForExport

                myWaiver.WaiverId = waivers(COLUMN_INDEX_WAIVERID, i)
                myWaiver.FirstName = waivers(COLUMN_INDEX_FIRSTNAME, i)
                myWaiver.LastName = waivers(COLUMN_INDEX_LASTNAME, i)
                myWaiver.Age = waivers(COLUMN_INDEX_AGE, i)
                myWaiver.DateOfBirth = waivers(COLUMN_INDEX_DATEOFBIRTH, i)
                myWaiver.PhoneNumber = waivers(COLUMN_INDEX_PHONENUMBER, i)
                myWaiver.Address = waivers(COLUMN_INDEX_ADDRESS, i)
                myWaiver.City = waivers(COLUMN_INDEX_CITY, i)
                myWaiver.State = waivers(COLUMN_INDEX_STATE, i)
                myWaiver.ZipCode = waivers(COLUMN_INDEX_ZIPCODE, i)
                myWaiver.PlayDate = waivers(COLUMN_INDEX_PLAYDATE, i)
                myWaiver.CreateDateTime = waivers(COLUMN_INDEX_WAIVERDATE, i)
                myWaiver.EmailAddress = waivers(COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.EmailList = waivers(COLUMN_INDEX_EMAILLIST, i)
                myWaiver.EventId = waivers(COLUMN_INDEX_EVENTID, i)
                myWaiver.HashId = waivers(COLUMN_INDEX_HASHID, i)
                myWaiver.ParentFirstName = waivers(COLUMN_INDEX_PARENTFIRSTNAME, i)
                myWaiver.ParentLastName = waivers(COLUMN_INDEX_PARENTLASTNAME, i)
                myWaiver.PlayTime = waivers(COLUMN_INDEX_PLAYTIME, i)
                myWaiver.SubmittedFromIpAddress = waivers(COLUMN_INDEX_SUBMITTEDFROMIPADDRESS, i)
                myWaiver.WaiverCustomFields = waivers(COLUMN_INDEX_CUSTOMFIELDS, i)

                mo_List.Add myWaiver
            next
        end if

        mn_ExportOption = EXPORTOPTION_YEAR
    end sub


    public sub LoadUngroupedByDateAndTime(dtPlayDate, dtPlayTime, IncludeInvalid, IncludeCheckedIn)
        dim myCon
        set myCon = new WaiverDataConnection

        dim waivers
        waivers = myCon.GetWaiversDateAndTimeWithoutEvent(dtPlayDate, dtPlayTime, IncludeInvalid, IncludeCheckedIn)

        if UBound(waivers) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(waivers, 2) 
                dim myWaiver
                set myWaiver = new Waiver

                myWaiver.WaiverId = waivers(W2_COLUMN_INDEX_WAIVERID, i)
                myWaiver.CreateDateTime = waivers(W2_COLUMN_INDEX_CREATEDATETIME, i)
                myWaiver.PlayerId = waivers(W2_COLUMN_INDEX_PLAYERID, i)
                myWaiver.ParentId = waivers(W2_COLUMN_INDEX_PARENTID, i)
                myWaiver.PhoneNumber = waivers(W2_COLUMN_INDEX_PHONENUMBER, i)
                myWaiver.Address = waivers(W2_COLUMN_INDEX_ADDRESS, i)
                myWaiver.City = waivers(W2_COLUMN_INDEX_CITY, i)
                myWaiver.State = waivers(W2_COLUMN_INDEX_STATE, i)
                myWaiver.ZipCode = waivers(W2_COLUMN_INDEX_ZIPCODE, i)
                myWaiver.EmailAddress = waivers(W2_COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.LegaleseId = waivers(W2_COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.HashId = waivers(W2_COLUMN_INDEX_HASHID, i)
                myWaiver.SubmittedFromIpAddress = waivers(W2_COLUMN_INDEX_SUBMITTEDFROMIPADDRESS, i)
                myWaiver.Signature = waivers(W2_COLUMN_INDEX_SIGNATURE, i)
                myWaiver.Exported = waivers(W2_COLUMN_INDEX_EXPORTED, i) 

                mo_List.Add myWaiver
            next
        end if
    end sub

    public sub LoadUngroupedByDate(dtPlayDate, IncludeInvalid, IncludeCheckedIn)
        dim myCon
        set myCon = new WaiverDataConnection

        dim waivers
        waivers = myCon.GetWaiversForDateWithoutEvent(dtPlayDate, IncludeInvalid, IncludeCheckedIn)

        if UBound(waivers) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(waivers, 2) 
                dim myWaiver
                set myWaiver = new Waiver

                myWaiver.WaiverId = waivers(W2_COLUMN_INDEX_WAIVERID, i)
                myWaiver.CreateDateTime = waivers(W2_COLUMN_INDEX_CREATEDATETIME, i)
                myWaiver.PlayerId = waivers(W2_COLUMN_INDEX_PLAYERID, i)
                myWaiver.ParentId = waivers(W2_COLUMN_INDEX_PARENTID, i)
                myWaiver.PhoneNumber = waivers(W2_COLUMN_INDEX_PHONENUMBER, i)
                myWaiver.Address = waivers(W2_COLUMN_INDEX_ADDRESS, i)
                myWaiver.City = waivers(W2_COLUMN_INDEX_CITY, i)
                myWaiver.State = waivers(W2_COLUMN_INDEX_STATE, i)
                myWaiver.ZipCode = waivers(W2_COLUMN_INDEX_ZIPCODE, i)
                myWaiver.EmailAddress = waivers(W2_COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.LegaleseId = waivers(W2_COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.HashId = waivers(W2_COLUMN_INDEX_HASHID, i)
                myWaiver.SubmittedFromIpAddress = waivers(W2_COLUMN_INDEX_SUBMITTEDFROMIPADDRESS, i)
                myWaiver.Signature = waivers(W2_COLUMN_INDEX_SIGNATURE, i)
                myWaiver.Exported = waivers(W2_COLUMN_INDEX_EXPORTED, i) 

                mo_List.Add myWaiver
            next
        end if
    end sub

    public sub LoadWaiversByPlayerId(nPlayerId)
        dim myCon
        set myCon = new WaiverDataConnection

        dim waivers
        waivers = myCon.GetWaiversByPlayerId(nPlayerId)

        if UBound(waivers) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(waivers, 2) 
                dim myWaiver
                set myWaiver = new Waiver

                myWaiver.WaiverId = waivers(W2_COLUMN_INDEX_WAIVERID, i)
                myWaiver.CreateDateTime = waivers(W2_COLUMN_INDEX_CREATEDATETIME, i)
                myWaiver.PlayerId = waivers(W2_COLUMN_INDEX_PLAYERID, i)
                myWaiver.ParentId = waivers(W2_COLUMN_INDEX_PARENTID, i)
                myWaiver.PhoneNumber = waivers(W2_COLUMN_INDEX_PHONENUMBER, i)
                myWaiver.Address = waivers(W2_COLUMN_INDEX_ADDRESS, i)
                myWaiver.City = waivers(W2_COLUMN_INDEX_CITY, i)
                myWaiver.State = waivers(W2_COLUMN_INDEX_STATE, i)
                myWaiver.ZipCode = waivers(W2_COLUMN_INDEX_ZIPCODE, i)
                myWaiver.EmailAddress = waivers(W2_COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.LegaleseId = waivers(W2_COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.HashId = waivers(W2_COLUMN_INDEX_HASHID, i)
                myWaiver.SubmittedFromIpAddress = waivers(W2_COLUMN_INDEX_SUBMITTEDFROMIPADDRESS, i)
                myWaiver.Signature = waivers(W2_COLUMN_INDEX_SIGNATURE, i)
                myWaiver.Exported = waivers(W2_COLUMN_INDEX_EXPORTED, i)               

                mo_List.Add myWaiver
            next
        end if
    end sub

    public sub LoadWaiversByEmail(sEmailAddress)
        dim myCon
        set myCon = new WaiverDataConnection

        dim waivers
        waivers = myCon.GetWaiversByEmail(sEmailAddress)

        if UBound(waivers) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(waivers, 2) 
                dim myWaiver
                set myWaiver = new Waiver

                myWaiver.WaiverId = waivers(W2_COLUMN_INDEX_WAIVERID, i)
                myWaiver.CreateDateTime = waivers(W2_COLUMN_INDEX_CREATEDATETIME, i)
                myWaiver.PlayerId = waivers(W2_COLUMN_INDEX_PLAYERID, i)
                myWaiver.ParentId = waivers(W2_COLUMN_INDEX_PARENTID, i)
                myWaiver.PhoneNumber = waivers(W2_COLUMN_INDEX_PHONENUMBER, i)
                myWaiver.Address = waivers(W2_COLUMN_INDEX_ADDRESS, i)
                myWaiver.City = waivers(W2_COLUMN_INDEX_CITY, i)
                myWaiver.State = waivers(W2_COLUMN_INDEX_STATE, i)
                myWaiver.ZipCode = waivers(W2_COLUMN_INDEX_ZIPCODE, i)
                myWaiver.EmailAddress = waivers(W2_COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.LegaleseId = waivers(W2_COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.HashId = waivers(W2_COLUMN_INDEX_HASHID, i)
                myWaiver.SubmittedFromIpAddress = waivers(W2_COLUMN_INDEX_SUBMITTEDFROMIPADDRESS, i)
                myWaiver.Signature = waivers(W2_COLUMN_INDEX_SIGNATURE, i)
                myWaiver.Exported = waivers(W2_COLUMN_INDEX_EXPORTED, i)               

                mo_List.Add myWaiver
            next
        end if
    end sub

    public sub LoadWaiversByPlayersInfo(FirstName, LastName, DOB)
        dim myCon
        set myCon = new WaiverDataConnection

        dim waivers
        waivers = myCon.GetWaiversByPlayersInfo(FirstName, LastName, DOB)

        if UBound(waivers) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(waivers, 2) 
                dim myWaiver
                set myWaiver = new Waiver

                myWaiver.WaiverId = waivers(W2_COLUMN_INDEX_WAIVERID, i)
                myWaiver.CreateDateTime = waivers(W2_COLUMN_INDEX_CREATEDATETIME, i)
                myWaiver.PlayerId = waivers(W2_COLUMN_INDEX_PLAYERID, i)
                myWaiver.ParentId = waivers(W2_COLUMN_INDEX_PARENTID, i)
                myWaiver.PhoneNumber = waivers(W2_COLUMN_INDEX_PHONENUMBER, i)
                myWaiver.Address = waivers(W2_COLUMN_INDEX_ADDRESS, i)
                myWaiver.City = waivers(W2_COLUMN_INDEX_CITY, i)
                myWaiver.State = waivers(W2_COLUMN_INDEX_STATE, i)
                myWaiver.ZipCode = waivers(W2_COLUMN_INDEX_ZIPCODE, i)
                myWaiver.EmailAddress = waivers(W2_COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.LegaleseId = waivers(W2_COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.HashId = waivers(W2_COLUMN_INDEX_HASHID, i)
                myWaiver.SubmittedFromIpAddress = waivers(W2_COLUMN_INDEX_SUBMITTEDFROMIPADDRESS, i)
                myWaiver.Signature = waivers(W2_COLUMN_INDEX_SIGNATURE, i)
                myWaiver.Exported = waivers(W2_COLUMN_INDEX_EXPORTED, i)               

                mo_List.Add myWaiver
            next
        end if
    end sub

    public function MarkAsExported()
        dim myCon
        set myCon = new WaiverDataConnection

        select case mn_ExportOption
            case EXPORTOPTION_ALL
                myCon.UpdateWaiverExportForAll
            case EXPORTOPTION_CUSTOM
                myCon.UpdateWaiverExportForCustomRange mdt_ExportStart, mdt_ExportEnd
            case EXPORTOPTION_MONTH
                myCon.UpdateWaiverExportForMonth
            case EXPORTOPTION_YEAR
                myCon.UpdateWaiverExportForYear
        end select
    end function

    public sub Search(sText, IncludeInvalid, IncludeCheckedIn)
        dim myCon
        set myCon = new WaiverDataConnection

        dim waivers
        waivers = myCon.SearchWaivers(sText, IncludeInvalid, IncludeCheckedIn)

        if UBound(waivers) > 0 then
            ' using the GetRows method of the ADODB.RecordSet puts the resulting array in Column, Row format
            for i = 0 to ubound(waivers, 2) 
                dim myWaiver
                set myWaiver = new Waiver

                myWaiver.WaiverId = waivers(W2_COLUMN_INDEX_WAIVERID, i)
                myWaiver.CreateDateTime = waivers(W2_COLUMN_INDEX_CREATEDATETIME, i)
                myWaiver.PlayerId = waivers(W2_COLUMN_INDEX_PLAYERID, i)
                myWaiver.ParentId = waivers(W2_COLUMN_INDEX_PARENTID, i)
                myWaiver.PhoneNumber = waivers(W2_COLUMN_INDEX_PHONENUMBER, i)
                myWaiver.Address = waivers(W2_COLUMN_INDEX_ADDRESS, i)
                myWaiver.City = waivers(W2_COLUMN_INDEX_CITY, i)
                myWaiver.State = waivers(W2_COLUMN_INDEX_STATE, i)
                myWaiver.ZipCode = waivers(W2_COLUMN_INDEX_ZIPCODE, i)
                myWaiver.EmailAddress = waivers(W2_COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.LegaleseId = waivers(W2_COLUMN_INDEX_EMAILADDRESS, i)
                myWaiver.HashId = waivers(W2_COLUMN_INDEX_HASHID, i)
                myWaiver.SubmittedFromIpAddress = waivers(W2_COLUMN_INDEX_SUBMITTEDFROMIPADDRESS, i)
                myWaiver.Signature = waivers(W2_COLUMN_INDEX_SIGNATURE, i)
                myWaiver.Exported = waivers(W2_COLUMN_INDEX_EXPORTED, i)               

                mo_List.Add myWaiver
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