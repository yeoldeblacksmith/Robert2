<%
class UserDataConnection
    public sub ChangePassword(UserName, Password)
        dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "ChangePassword"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@UserName", adVarChar, adParamInput, 30, UserName)
        oCmd.Parameters.Append oCmd.CreateParameter("@Password", adVarChar, adParamInput, 50, Password)

        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub

    public sub DeleteUser(UserName)
        dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "DeleteUser"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@UserName", adVarChar, adParamInput, 30, UserName)

        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub

    public function GetAllUsersForSite()
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetAllUsersForSite"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        
        oCon.GetRecordSet oCmd, oRs

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

        GetAllUsersForSite = results
    end function

    public function GetUser(UserName)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetUser"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@UserName", adVarChar, adParamInput, 30, UserName)
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        
        oCon.GetRecordSet oCmd, oRs

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

        GetUser = results
    end function

    public function GetUserByEmail(Email)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetUserByEmail"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@EmailAddress", adVarChar, adParamInput, 50, Email)
        
        oCon.GetRecordSet oCmd, oRs

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

        GetUserByEmail = results
    end function

    public function GetUserByHash(PasswordResetHash)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetUserByHash"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@PasswordResetHash", adVarChar, adParamInput, 50, PasswordResetHash)
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
        
        oCon.GetRecordSet oCmd, oRs

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

        GetUserByHash = results
    end function

    public sub SavePasswordResetHash(UserName, PasswordResetHash, SiteGuid)
        dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "SavePasswordResetHash"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@UserName", adVarChar, adParamInput, 30, UserName)
        oCmd.Parameters.Append oCmd.CreateParameter("@PasswordResetHash", adVarChar, adParamInput, 50, PasswordResetHash)
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, SiteGuid)

        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub

    public sub SaveUser(UserName, RoleId, EmailAddress, Enabled, PasswordExpired, SiteGuid)
        dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "SaveUser"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@UserName", adVarChar, adParamInput, 30, UserName)
        oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, SiteGuid)
        oCmd.Parameters.Append oCmd.CreateParameter("@RoleId", adInteger, adParamInput, 4, RoleId)
        oCmd.Parameters.Append oCmd.CreateParameter("@EmailAddress", adVarChar, adParamInput, 50, EmailAddress)
        oCmd.Parameters.Append oCmd.CreateParameter("@Enabled", adBoolean, adParamInput, 1, Enabled)
        oCmd.Parameters.Append oCmd.CreateParameter("@PasswordExpired", adBoolean, adParamInput, 1, PasswordExpired)

        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub
end class
%>