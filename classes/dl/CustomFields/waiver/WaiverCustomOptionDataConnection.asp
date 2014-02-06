<%
    class WaiverCustomOptionDataConnection

        public sub DeleteWaiverCustomOption(WaiverCustomFieldId, Sequence)
            dim oCmd, oCon
            set oCmd = Server.CreateObject("ADODB.Command")
    
            oCmd.CommandText = "DeleteWaiverCustomOption"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@WaiverCustomFieldId", adInteger, adParamInput, 4, WaiverCustomFieldId)
            oCmd.Parameters.Append oCmd.CreateParameter("@Sequence", adInteger, adParamInput, 4, Sequence)
        
            set oCon = new DataConnection
            oCon.ExecuteCommand oCmd

            set oCmd = nothing
            oCon.CloseConnection
        end sub

        public function GetAllActiveWaiverCustomOptionsByFieldId(WaiverCustomFieldId)
            dim oCmd, oRs, oCon

            set oCmd = Server.CreateObject("ADODB.Command")
            set oRs = Server.CreateObject("ADODB.RecordSet")
            set oCon = new DataConnection

            oCmd.CommandText = "GetAllActiveWaiverCustomOptionsByFieldId"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@WaiverCustomFieldId", adInteger, adParamInput, 4, WaiverCustomFieldId)

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

            GetAllActiveWaiverCustomOptionsByFieldId = results
        end function

        public function GetNextWaiverCustomOptionSequence(WaiverCustomFieldId)
            dim oCmd, oCon
            set oCmd = Server.CreateObject("ADODB.Command")

            oCmd.CommandText = "GetNextWaiverCustomOptionSequence"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@Return", adInteger, adParamReturnValue, 4,0)
            oCmd.Parameters.Append oCmd.CreateParameter("@WaiverCustomFieldId", adInteger, adParamInput, 4, WaiverCustomFieldId)

            set oCon = new DataConnection
            oCon.ExecuteCommand oCmd

            GetNextWaiverCustomOptionSequence = oCmd.Parameters("@Return").Value

            oCon.CloseConnection

            set oCmd = nothing
        end function

        public function GetWaiverCustomOption(WaiverCustomFieldId, Sequence)
            dim oCmd, oRs, oCon

            set oCmd = Server.CreateObject("ADODB.Command")
            set oRs = Server.CreateObject("ADODB.RecordSet")
            set oCon = new DataConnection

            oCmd.CommandText = "GetWaiverCustomOption"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@WaiverCustomFieldId", adInteger, adParamInput, 4, WaiverCustomFieldId)
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

            GetWaiverCustomOption = results
        end function

        public sub SaveWaiverCustomOption(WaiverCustomFieldId, Sequence, Text, Value)
            dim oCmd, oCon
            set oCmd = Server.CreateObject("ADODB.Command")

            oCmd.CommandText = "SaveWaiverCustomOption"
            oCmd.CommandType = adCmdStoredProc

            ' required parameters
            oCmd.Parameters.Append oCmd.CreateParameter("@WaiverCustomFieldId", adInteger, adParamInput, 4, WaiverCustomFieldId)
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