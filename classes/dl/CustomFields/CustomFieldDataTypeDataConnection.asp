<%
    class CustomFieldDataTypeDataConnection

        public function GetAllCustomFieldDataTypes()
            dim oCmd, oRs, oCon

            set oCmd = Server.CreateObject("ADODB.Command")
            set oRs = Server.CreateObject("ADODB.RecordSet")
            set oCon = new DataConnection

            oCmd.CommandText = "GetAllCustomFieldDataTypes"
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

            GetAllCustomFieldDataTypes = results
        end function

        public function GetCustomFieldDataType(CustomFieldDataTypeId)
            dim oCmd, oRs, oCon

            set oCmd = Server.CreateObject("ADODB.Command")
            set oRs = Server.CreateObject("ADODB.RecordSet")
            set oCon = new DataConnection

            oCmd.CommandText = "GetCustomFieldDataType"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@CustomFieldDataTypeId", adInteger, adParamInput, 4, CustomFieldDataTypeId)

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

            GetCustomFieldDataType = results
        end function

    end class
%>