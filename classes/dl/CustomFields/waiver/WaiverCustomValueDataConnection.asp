<%
    class WaiverCustomValueDataConnection

        public sub DeleteWaiverCustomValueByWaiverId(WaiverId)
            dim oCmd, oCon
            set oCmd = Server.CreateObject("ADODB.Command")

            oCmd.CommandText = "DeleteWaiverCustomValueByWaiverId"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@WaiverId", adInteger, adParamInput, 4, WaiverId)
        
            set oCon = new DataConnection
            oCon.ExecuteCommand oCmd

            set oCmd = nothing
            oCon.CloseConnection
        end sub

        public function GetAllWaiverCustomValuesByWaiverId(WaiverId)
            dim oCmd, oRs, oCon

            set oCmd = Server.CreateObject("ADODB.Command")
            set oRs = Server.CreateObject("ADODB.RecordSet")
            set oCon = new DataConnection

            oCmd.CommandText = "GetAllWaiverCustomValuesByWaiverId"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@WaiverId", adInteger, adParamInput, 4, WaiverId)

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

            GetAllWaiverCustomValuesByWaiverId = results
        end function

        public function GetAllActiveWaiverCustomValuesByWaiverId(WaiverId)
            dim oCmd, oRs, oCon

            set oCmd = Server.CreateObject("ADODB.Command")
            set oRs = Server.CreateObject("ADODB.RecordSet")
            set oCon = new DataConnection

            oCmd.CommandText = "GetAllActiveWaiverCustomValuesByWaiverId"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@WaiverId", adInteger, adParamInput, 4, WaiverId)

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

            GetAllActiveWaiverCustomValuesByWaiverId = results
        end function

        public sub SaveWaiverCustomValue(WaiverId, WaiverCustomFieldId, Value)
            dim oCmd, oCon
            set oCmd = Server.CreateObject("ADODB.Command")

            oCmd.CommandText = "SaveWaiverCustomValue"
            oCmd.CommandType = adCmdStoredProc

            ' required parameters
            oCmd.Parameters.Append oCmd.CreateParameter("@WaiverId", adInteger, adParamInput, 4, WaiverId)
            oCmd.Parameters.Append oCmd.CreateParameter("@WaiverCustomFieldId", adInteger, adParamInput, 4, WaiverCustomFieldId)
        
            ' optional parameters
            if len(Value) = 0 then
                oCmd.Parameters.Append oCmd.CreateParameter("@Value", adVarChar, adParamInput, 100, null)
            else
                oCmd.Parameters.Append oCmd.CreateParameter("@Value", adVarChar, adParamInput, 100, Value)
            end if

            set oCon = new DataConnection
            oCon.ExecuteCommand oCmd

            set oCmd = nothing
            oCon.CloseConnection
        end sub

    end class
%>