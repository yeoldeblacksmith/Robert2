<%
    class RegistrationCustomOptionDataConnection

        public sub DeleteRegistrationCustomOption(RegistrationCustomFieldId, Sequence)
            dim oCmd, oCon
            set oCmd = Server.CreateObject("ADODB.Command")

            oCmd.CommandText = "DeleteRegistrationCustomOption"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@RegistrationCustomFieldId", adInteger, adParamInput, 4, RegistrationCustomFieldId)
            oCmd.Parameters.Append oCmd.CreateParameter("@Sequence", adInteger, adParamInput, 4, Sequence)
        
            set oCon = new DataConnection
            oCon.ExecuteCommand oCmd

            set oCmd = nothing
            oCon.CloseConnection
        end sub

        public function GetAllActiveRegistrationCustomOptionsByFieldId(RegistrationCustomFieldId)
            dim oCmd, oRs, oCon

            set oCmd = Server.CreateObject("ADODB.Command")
            set oRs = Server.CreateObject("ADODB.RecordSet")
            set oCon = new DataConnection

            oCmd.CommandText = "GetAllActiveRegistrationCustomOptionsByFieldId"
            oCmd.CommandType = adCmdStoredProc
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

            GetAllActiveRegistrationCustomOptionsByFieldId = results
        end function

        public function GetNextRegistrationCustomOptionSequence(RegistrationCustomFieldId)
            dim oCmd, oCon
            set oCmd = Server.CreateObject("ADODB.Command")

            oCmd.CommandText = "GetNextRegistrationCustomOptionSequence"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@Return", adInteger, adParamReturnValue, 4,0)
            oCmd.Parameters.Append oCmd.CreateParameter("@RegistrationCustomFieldId", adInteger, adParamInput, 4, RegistrationCustomFieldId)

            set oCon = new DataConnection
            oCon.ExecuteCommand oCmd

            GetNextRegistrationCustomOptionSequence = oCmd.Parameters("@Return").Value

            oCon.CloseConnection

            set oCmd = nothing
        end function
        
        public function GetRegistrationCustomOption(RegistrationCustomFieldId, Sequence)
            dim oCmd, oRs, oCon

            set oCmd = Server.CreateObject("ADODB.Command")
            set oRs = Server.CreateObject("ADODB.RecordSet")
            set oCon = new DataConnection

            oCmd.CommandText = "GetRegistrationCustomOption"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@RegistrationCustomFieldId", adInteger, adParamInput, 4, RegistrationCustomFieldId)
            oCmd.Parameters.Append oCmd.CreateParameter("@Sequence", adInteger, adParamInput, 4, Sequence)

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

            GetRegistrationCustomOption = results
        end function

        public sub SaveRegistrationCustomOption(RegistrationgCustomFieldId, Sequence, Text, Value)
            dim oCmd, oCon
            set oCmd = Server.CreateObject("ADODB.Command")

            oCmd.CommandText = "SaveRegistrationCustomOption"
            oCmd.CommandType = adCmdStoredProc

            ' required parameters
            oCmd.Parameters.Append oCmd.CreateParameter("@RegistrationCustomFieldId", adInteger, adParamInput, 4, RegistrationgCustomFieldId)
            oCmd.Parameters.Append oCmd.CreateParameter("@Sequence", adInteger, adParamInput, 4, Sequence)
        
            ' optional parameters
            'if len(Text) = 0 then
            '    oCmd.Parameters.Append oCmd.CreateParameter("@Text", adVarChar, adParamInput, 100, null)
            'else
                oCmd.Parameters.Append oCmd.CreateParameter("@Text", adVarChar, adParamInput, 100, Text)
            'end if
            'if len(Value) = 0 then
            '    oCmd.Parameters.Append oCmd.CreateParameter("@Value", adVarChar, adParamInput, 100, null)
            'else
                oCmd.Parameters.Append oCmd.CreateParameter("@Value", adVarChar, adParamInput, 100, Value)
            'end if

            set oCon = new DataConnection
            oCon.ExecuteCommand oCmd

            set oCmd = nothing
            oCon.CloseConnection
        end sub

    end class
%>