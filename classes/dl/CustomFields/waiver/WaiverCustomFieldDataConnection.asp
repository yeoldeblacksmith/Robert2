<%
    class WaiverCustomFieldDataConnection
        
        public sub DeleteWaiverCustomField(WaiverCustomFieldId)
            dim oCmd, oCon
            set oCmd = Server.CreateObject("ADODB.Command")

            oCmd.CommandText = "DeleteWaiverCustomField"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@WaiverCustomFieldId", adInteger, adParamInput, 4, WaiverCustomFieldId)
        
            set oCon = new DataConnection
            oCon.ExecuteCommand oCmd

            set oCmd = nothing
            oCon.CloseConnection
        end sub

        public function GetAllActiveWaiverCustomFields()
            dim oCmd, oRs, oCon

            set oCmd = Server.CreateObject("ADODB.Command")
            set oRs = Server.CreateObject("ADODB.RecordSet")
            set oCon = new DataConnection

            oCmd.CommandText = "GetAllActiveWaiverCustomFields"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)

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

            GetAllActiveWaiverCustomFields = results
        end function

        public function GetDistinctCustomFields(includeInActive)
            dim oCmd, oRs, oCon

            set oCmd = Server.CreateObject("ADODB.Command")
            set oRs = Server.CreateObject("ADODB.RecordSet")
            set oCon = new DataConnection

            oCmd.CommandText = "GetDistinctCustomFields"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
            oCmd.Parameters.Append oCmd.CreateParameter("@IncludeInActive", adBoolean, adParamInput, 1, includeInActive)

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

            GetDistinctCustomFields = results
        end function
       

        public function GetWaiverCustomField(WaiverCustomFieldId)
            dim oCmd, oRs, oCon

            set oCmd = Server.CreateObject("ADODB.Command")
            set oRs = Server.CreateObject("ADODB.RecordSet")
            set oCon = new DataConnection

            oCmd.CommandText = "GetWaiverCustomField"
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

            GetWaiverCustomField = results
        end function

        public sub SaveWaiverCustomField(WaiverCustomFieldId, Name, CustomFieldDataTypeId, Notes, Required, ValidateDateType, DefaultValue, FormLocation, AdditionalInformation)
    'response.Write "DEBUG: WaiverCustomFieldId = " & WaiverCustomFieldId & "<br/>"
    'response.Write "DEBUG: Name = " & Name & "<br/>"
    'response.Write "DEBUG: CustomFieldDataTypeId = " & CustomFieldDataTypeId & "<br/>"
    'response.Write "DEBUG: Notes = " & Notes & "<br/>"
    'response.Write "DEBUG: Required = " & Required & "<br/>"
    'response.Write "DEBUG: ValidateDataType = " & ValidateDataType & "<br/>"
    'response.Write "DEBUG: DefaultValue = " & DefaultValue & "<br/>"

            dim oCmd, oCon
            set oCmd = Server.CreateObject("ADODB.Command")

            oCmd.CommandText = "SaveWaiverCustomField"
            oCmd.CommandType = adCmdStoredProc

            ' required parameters
            oCmd.Parameters.Append oCmd.CreateParameter("@WaiverCustomFieldId", adInteger, adParamInput, 4, WaiverCustomFieldId)
            oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
            oCmd.Parameters.Append oCmd.CreateParameter("@Name", adVarChar, adParamInput, 100, Name)
            oCmd.Parameters.Append oCmd.CreateParameter("@CustomFieldDataTypeId", adInteger, adParamInput, 4, CustomFieldDataTypeId)
            
            ' optional parameters
            if len(Notes) = 0 then
                oCmd.Parameters.Append oCmd.CreateParameter("@Notes", adVarChar, adParamInput, -1, null)
            else
                oCmd.Parameters.Append oCmd.CreateParameter("@Notes", adVarChar, adParamInput, -1, Notes)
            end if
            if len(Required) = 0 then
                oCmd.Parameters.Append oCmd.CreateParameter("@Required", adBoolean, adParamInput, 1, null)
            else
                oCmd.Parameters.Append oCmd.CreateParameter("@Required", adBoolean, adParamInput, 1, Required)
            end if
            if len(ValidateDataType) = 0 then
                oCmd.Parameters.Append oCmd.CreateParameter("@ValidateDataType", adBoolean, adParamInput, 1, null)
            else
                oCmd.Parameters.Append oCmd.CreateParameter("@ValidateDataType", adBoolean, adParamInput, 1, ValidateDataType)
            end if
            if len(DefaultValue) = 0 then
                oCmd.Parameters.Append oCmd.CreateParameter("@DefaultValue", adVarChar, adParamInput, -1, null)
            else
                oCmd.Parameters.Append oCmd.CreateParameter("@DefaultValue", adVarChar, adParamInput, -1, DefaultValue)
            end if
            if len(FormLocation) = 0 then
                oCmd.Parameters.Append oCmd.CreateParameter("@FormLocation", adInteger, adParamInput, -1, WAIVERCUSTOMFIELDS_LOCATION_WAIVERGENERAL)
            else
                oCmd.Parameters.Append oCmd.CreateParameter("@FormLocation", adInteger, adParamInput, -1, FormLocation)
            end if

            if len(AdditionalInformation) = 0 then
                oCmd.Parameters.Append oCmd.CreateParameter("@AdditionalInformation", adVarChar, adParamInput, -1, null)
            else
                oCmd.Parameters.Append oCmd.CreateParameter("@AdditionalInformation", adVarChar, adParamInput, -1, AdditionalInformation)
            end if

            set oCon = new DataConnection
            oCon.ExecuteCommand oCmd

            set oCmd = nothing
            oCon.CloseConnection
        end sub

    end class
%>