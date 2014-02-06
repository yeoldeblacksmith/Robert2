<%
class RoleDataConnection
    function GetAllRoles()
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.RecordSet")
        set oCon = new DataConnection

        oCmd.CommandText = "GetAllRoles"
        oCmd.CommandType = adCmdStoredProc

        oCon.GetRecordSet oCmd, oRs

        dim results
        if oRs.EOF then
            results = Array()
        else
            results = oRs.GetRows()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetAllRoles = results
    end function

    function GetRole(RoleId)
        dim oCmd, oRs, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.RecordSet")
        set oCon = new DataConnection

        oCmd.CommandText = "GetRole"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@RoleId", adInteger, adParamInput, 4, RoleId)

        oCon.GetRecordSet oCmd, oRs

        dim results
        if oRs.EOF then
            results = Array()
        else
            results = oRs.GetRows()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetRole = results
    end function

end class
%>