<%
    class RegistrationCustomField
        ' attributes
        dim mn_Id, ms_SiteGuid, ms_Name, _
            mo_DataType, ms_Notes, mb_Required, _
            mb_Validate, ms_DefaultValue, mo_Options, _
            ms_AdditionalInfo, mn_Sequence, mn_PaymentTypeId

        ' ctor
        public sub Class_Initialize()
            set mo_DataType = new CustomDataType
            set mo_Options = new RegistrationCustomOptionCollection

            ms_Notes = ""
            ms_DefaultValue = ""
            ms_AdditionalInfo = ""
        end sub

        ' methods
        public sub Delete()
            dim myCon
            set myCon = new RegistrationCustomFieldDataConnection

            myCon.DeleteRegistrationCustomField RegistrationCustomFieldId
        end sub

        public function GetFormValue(EventId)
            dim myValue
            set myValue = new RegistrationCustomValue

            myValue.Load EventId, RegistrationCustomFieldId

            if isnull(myValue) then
                GetFormValue = ""
            else
                GetFormValue = myValue.Value
            end if
        end function

        public sub Load()
            dim myCon
            set myCon = new RegistrationCustomFieldDataConnection

            dim results
            results = myCon.GetRegistrationCustomField(RegistrationCustomFieldId)

            if ubound(results) > 0 then
                RegistrationCustomFieldId = results(REGISTRATIONCUSTOMFIELD_INDEX_ID, 0)
                SiteGuid = results(REGISTRATIONCUSTOMFIELD_INDEX_SITEGUID, 0)
                Name = results(REGISTRATIONCUSTOMFIELD_INDEX_NAME, 0)
                CustomFieldDataType.CustomFieldDataTypeId = results(REGISTRATIONCUSTOMFIELD_INDEX_DATATYPE, 0)
                Notes = results(REGISTRATIONCUSTOMFIELD_INDEX_NOTES, 0)
                Required = results(REGISTRATIONCUSTOMFIELD_INDEX_REQUIRED, 0)
                ValidateDataType = results(REGISTRATIONCUSTOMFIELD_INDEX_VALIDATE, 0)
                DefaultValue = results(REGISTRATIONCUSTOMFIELD_INDEX_DEFAULTVALUE, 0)
                AdditionalInformation = results(REGISTRATIONCUSTOMFIELD_INDEX_ADDITIONALINFO, 0)
                Sequence = results(REGISTRATIONCUSTOMFIELD_INDEX_SEQUENCE, 0)
                PaymentTypeId = results(REGISTRATIONCUSTOMFIELD_INDEX_PAYMENTTYPE, 0)
            end if
        end sub
    
        public sub LoadById(nId)
            RegistrationCustomFieldId = nId

            Load
        end sub    

        public sub LoadOptions()
            Options.LoadByFieldId RegistrationCustomFieldId
        end sub

        public sub Save()
            dim myCon
            set myCon = new RegistrationCustomFieldDataConnection

            'myCon.SaveRegistrationCustomField RegistrationCustomFieldId, Name, mo_DataType.CustomFieldDataTypeId, Notes, Required, ValidateDataType, DefaultValue
            myCon.SaveRegistrationCustomField RegistrationCustomFieldId, Name, CustomFieldDataType.CustomFieldDataTypeId, _
                                              Notes, Required, ValidateDataType, _
                                              DefaultValue, AdditionalInformation, Sequence, _
                                              PaymentTypeId
        end sub

        public sub SaveFormValue(EventId, Value)
            dim myValue
            set myValue = new RegistrationCustomValue

            myValue.EventId = EventId
            myValue.RegistrationCustomFieldId = RegistrationCustomFieldId

            if CustomFieldDataType.CustomFieldDataTypeId = FIELDTYPE_YESNO then
                myValue.Value = iif(Value = "on", "true", "false")
            else
                myValue.Value = Value
            end if

            myValue.Save
        end sub

        ' properties
        public property Get RegistrationCustomFieldId
            RegistrationCustomFieldId = mn_Id
        end property

        public property Let RegistrationCustomFieldId(value)
            mn_Id = value
        end property

        public property Get SiteGuid
            SiteGuid = ms_SiteGuid
        end property

        public property Let SiteGuid(value)
            ms_SiteGuid = value
        end property

        public property Get Name  
            Name = ms_Name
        end property

        public property Let Name(value)
            ms_Name = value
        end property

        public property Get CustomFieldDataType
            set CustomFieldDataType = mo_DataType
        end property

        public property Let CustomFieldDataType(value)
            mo_DataType = value
        end property

        public property Get Notes
            Notes = ms_Notes
        end property

        public property Let Notes(value)
            if IsNull(value) then value = ""
            ms_Notes = value
        end property

        public property Get Required
            Required = mb_Required
        end property

        public property Let Required(value)
            mb_Required = value
        end property

        public property Get ValidateDataType
            ValidateDataType = mb_Validate
        end property

        public property Let ValidateDataType(Value)
            mb_Validate = value
        end property

        public property Get DefaultValue
            if IsNull(value) then value = ""
            DefaultValue = ms_DefaultValue
        end property

        public property Let DefaultValue(value)
            ms_DefaultValue = value
        end property
  
        public property get Options
            set Options = mo_Options
        end property

        public property let Options(value)
            set mo_Options = value
        end property

        public property Get AdditionalInformation
            AdditionalInformation = ms_AdditionalInfo
        end property

        public property Let AdditionalInformation(value)
            if IsNull(value) then value = ""
            ms_AdditionalInfo = value
        end property

        public property get Sequence
            Sequence = mn_Sequence
        end property

        public property let Sequence(value)
            mn_Sequence = value
        end property

        public property get PaymentTypeId
            PaymentTypeId = mn_PaymentTypeId
        end property

        public property let PaymentTypeId(value)
            mn_PaymentTypeId = value
        end property

        public property get HasPayment
            HasPayment = cbool(cstr(PaymentTypeId) > PAYMENTTYPE_NONE)
        end property
    end class
%>