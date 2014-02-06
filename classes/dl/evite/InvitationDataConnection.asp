<%
class InvitationDataConnection
    public sub DeleteInvitation(InvitationId)
        dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "DeleteInvitation"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@InvitationId", adInteger, adParamInput, 4, InvitationId)

        oCon.ExecuteCommand oCmd
        oCon.CloseConnection

        set oCmd = nothing
    end sub


    public function GetAllInvitationsByEventId(EventId)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetAllInvitationsByEventId"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)

        set oCon = new DataConnection
        set oRs = Server.CreateObject("ADODB.Recordset")
        oCon.GetRecordset oCmd, oRs

        dim results
        if oRs.EOF = false then
            results = oRs.GetRows()
        else
            results = Array()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetAllInvitationsByEventId = results
    end function
    
    public function GetAllInvitationsUnsentByEventId(EventId)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetAllInvitationsUnsentByEventId"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)

        set oCon = new DataConnection
        set oRs = Server.CreateObject("ADODB.Recordset")
        oCon.GetRecordset oCmd, oRs

        dim results
        if oRs.EOF = false then
            results = oRs.GetRows()
        else
            results = Array()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetAllInvitationsUnsentByEventId = results
    end function
    
    public function GetInvitation(InvitationId)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetInvitation"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@InvitationId", adInteger, adParamInput, 4, InvitationId)

        set oCon = new DataConnection
        set oRs = Server.CreateObject("ADODB.Recordset")
        oCon.GetRecordset oCmd, oRs

        dim results
        if oRs.EOF = false then
            results = oRs.GetRows()
        else
            results = Array()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetInvitation = results
    end function

    public function SaveInvitation(InvitationId, EventId, Name, EmailAddress, InviteDateTime, RsvpCount, RsvpDateTime, RsvpStatus, Comments, PublicComments)
        dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "SaveInvitation"
        oCmd.CommandType = adCmdStoredProc
        
        ' required parameters
        oCmd.Parameters.Append oCmd.CreateParameter("@RetVal", adInteger, adParamReturnValue)
        oCmd.Parameters.Append oCmd.CreateParameter("@InvitationId", adInteger, adParamInput, 4, InvitationId)
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)
        
        'optional paramters
        if IsNull(Name) = false and len(Name) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@Name", adVarChar, adParamInput, 50, Name)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@Name", adVarChar, adParamInput, 50, null)
        end if
        if IsNull(EmailAddress) = false and len(EmailAddress) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@EmailAddress", adVarChar, adParamInput, 120, EmailAddress)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@EmailAddress", adVarChar, adParamInput, 120, null)
        end if
        if IsNull(InviteDateTime) = false and len(InviteDateTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@InviteDateTime", adDate, adParamInput, 8, InviteDateTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@InviteDateTime", adDate, adParamInput, 8, null)
        end if
        if IsNull(RsvpCount) = false and len(RsvpCount) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@RsvpCount", adVarChar, adParamInput, 10, RsvpCount)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@RsvpCount", adVarChar, adParamInput, 10, null)
        end if
        if IsNull(RsvpDateTime) = false and len(RsvpDateTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@RsvpDateTime", adDate, adParamInput, 8, RsvpDateTime)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@RsvpDateTime", adDate, adParamInput, 8, null)
        end if
        if IsNull(RsvpStatus) = false and len(RsvpStatus) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@RsvpStatus", adVarChar, adParamInput, 20, RsvpStatus)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@RsvpStatus", adVarChar, adParamInput, 20, null)
        end if
        if IsNull(Comments) = false and len(Comments) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@Comments", adVarChar, adParamInput, 1000, Comments)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@Comments", adVarChar, adParamInput, 1000, null)
        end if
        if IsNull(PublicComments) = false and len(PublicComments) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@RsvpCount", adBoolean, adParamInput, 1, CBool(PublicComments))
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@RsvpCount", adBoolean, adParamInput, 1, null)
        end if
        
        oCon.ExecuteCommand oCmd

        SaveInvitation = oCmd("@RetVal")

        oCon.CloseConnection


        set oCmd = nothing
    end function
end class
%>