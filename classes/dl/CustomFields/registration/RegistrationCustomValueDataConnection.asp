<%
    class RegistrationCustomValueDataConnection

        public function GetAllRegistrationCustomValuesByEventId(EventId)
            dim oCmd, oRs, oCon

            set oCmd = Server.CreateObject("ADODB.Command")
            set oRs = Server.CreateObject("ADODB.RecordSet")
            set oCon = new DataConnection

            oCmd.CommandText = "GetAllRegistrationCustomValuesByEventId"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)

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

            GetAllRegistrationCustomValuesByEventId = results
        end function

        public function GetRegistrationCustomValue(EventId, RegistrationCustomFieldId)
            dim oCmd, oRs, oCon

            set oCmd = Server.CreateObject("ADODB.Command")
            set oRs = Server.CreateObject("ADODB.RecordSet")
            set oCon = new DataConnection

            oCmd.CommandText = "GetRegistrationCustomValue"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)
            oCmd.Parameters.Append oCmd.CreateParameter("@RegistrationCustomFieldId", adInteger, adParamInput, 4, RegistrationCustomFieldId)

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

            GetRegistrationCustomValue = results
        end function

        public sub SaveRegistrationCustomValue(EventId, RegistrationCustomFieldId, Value)
            dim oCmd, oCon
            set oCmd = Server.CreateObject("ADODB.Command")

            oCmd.CommandText = "SaveRegistrationCustomValue"
            oCmd.CommandType = adCmdStoredProc

            ' required parameters
            oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)
            oCmd.Parameters.Append oCmd.CreateParameter("@RegistrationCustomFieldId", adInteger, adParamInput, 4, RegistrationCustomFieldId)
        
            ' optional parameters
            if len(Value) = 0 then
                oCmd.Parameters.Append oCmd.CreateParameter("@Value", adVarChar, adParamInput, 1000, null)
            else
                oCmd.Parameters.Append oCmd.CreateParameter("@Value", adVarChar, adParamInput, 1000, Value)
            end if

            set oCon = new DataConnection
            oCon.ExecuteCommand oCmd

            set oCmd = nothing
            oCon.CloseConnection
        end sub
    end class
%>