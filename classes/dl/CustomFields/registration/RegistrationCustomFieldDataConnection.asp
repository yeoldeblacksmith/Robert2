<%
    class RegistrationCustomFieldDataConnection

        public sub DeleteRegistrationCustomField(RegistrationCustomFieldId)
            dim oCmd, oCon
            set oCmd = Server.CreateObject("ADODB.Command")

            oCmd.CommandText = "DeleteRegistrationCustomField"
            oCmd.CommandType = adCmdStoredProc
            oCmd.Parameters.Append oCmd.CreateParameter("@RegistrationCustomFieldId", adInteger, adParamInput, 4, RegistrationCustomFieldId)
        
            set oCon = new DataConnection
            oCon.ExecuteCommand oCmd

            set oCmd = nothing
            oCon.CloseConnection
        end sub

        public function GetAllActiveRegistrationCustomFields()
            dim oCmd, oRs, oCon

            set oCmd = Server.CreateObject("ADODB.Command")
            set oRs = Server.CreateObject("ADODB.RecordSet")
            set oCon = new DataConnection

            oCmd.CommandText = "GetAllActiveRegistrationCustomFields"
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

            GetAllActiveRegistrationCustomFields = results
        end function

        public function GetRegistrationCustomField(RegistrationCustomFieldId)
            dim oCmd, oRs, oCon

            set oCmd = Server.CreateObject("ADODB.Command")
            set oRs = Server.CreateObject("ADODB.RecordSet")
            set oCon = new DataConnection

            oCmd.CommandText = "GetRegistrationCustomField"
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

            GetRegistrationCustomField = results
        end function

        public sub SaveRegistrationCustomField(RegistrationCustomerFieldId, Name, CustomFieldDataTypeId, _
                                               Notes, Required, ValidateDataType, _
                                               DefaultValue, AdditionalInformation, Sequence, _
                                               PaymentTypeId)
    'response.Write "DEBUG: RegistrationCustomerFieldId = " & RegistrationCustomerFieldId & "<br/>"
    'response.Write "DEBUG: Name = " & Name & "<br/>"
    'response.Write "DEBUG: CustomFieldDataTypeId = " & CustomFieldDataTypeId & "<br/>"
    'response.Write "DEBUG: Notes = " & Notes & "<br/>"
    'response.Write "DEBUG: Required = " & Required & "<br/>"
    'response.Write "DEBUG: ValidateDataType = " & ValidateDataType & "<br/>"
    'response.Write "DEBUG: DefaultValue = " & DefaultValue & "<br/>"

            dim oCmd, oCon
            set oCmd = Server.CreateObject("ADODB.Command")

            oCmd.CommandText = "SaveRegistrationCustomField"
            oCmd.CommandType = adCmdStoredProc

            ' required parameters
            oCmd.Parameters.Append oCmd.CreateParameter("@RegistrationCustomFieldId", adInteger, adParamInput, 4, RegistrationCustomerFieldId)
            oCmd.Parameters.Append oCmd.CreateParameter("@SiteGuid", adVarChar, adParamInput, 38, MY_SITE_GUID)
            oCmd.Parameters.Append oCmd.CreateParameter("@Name", adVarChar, adParamInput, 100, Name)
            oCmd.Parameters.Append oCmd.CreateParameter("@CustomFieldDataType", adInteger, adParamInput, 4, CustomFieldDataTypeId)
            
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
            if len(AdditionalInformation) = 0 then
                oCmd.Parameters.Append oCmd.CreateParameter("@AdditionalInformation", adVarChar, adParamInput, -1, null)
            else
                oCmd.Parameters.Append oCmd.CreateParameter("@AdditionalInformation", adVarChar, adParamInput, -1, AdditionalInformation)
            end if
            if len(Sequence) = 0 then
                oCmd.Parameters.Append oCmd.CreateParameter("@Sequence", adInteger, adParamInput, 4, null)
            else
                oCmd.Parameters.Append oCmd.CreateParameter("@Sequence", adInteger, adParamInput, 4, Sequence)
            end if
            if len(PaymentTypeId) = 0 then
                oCmd.Parameters.Append oCmd.CreateParameter("@Sequence", adUnsignedTinyInt, adParamInput, 2, null)
            else
                oCmd.Parameters.Append oCmd.CreateParameter("@Sequence", adUnsignedTinyInt, adParamInput, 2, PaymentTypeId)
            end if

            set oCon = new DataConnection
            oCon.ExecuteCommand oCmd

            set oCmd = nothing
            oCon.CloseConnection
        end sub

    end class
%>