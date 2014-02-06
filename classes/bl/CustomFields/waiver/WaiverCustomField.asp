<%
    class WaiverCustomField
        ' attributes
        dim mn_Id, ms_SiteGuid, ms_Name, _
            mo_DataType, ms_Notes, mb_Required, _
            mb_Validate, ms_DefaultValue, ms_FormLocation, ms_AdditionalInfo

        ' ctor
        public sub Class_Initialize()
            set mo_DataType = new CustomDataType
            ms_AdditionalInfo = ""
        end sub

        ' methods
        public sub Delete()
            dim myCon
            set myCon = new WaiverCustomFieldDataConnection

            myCon.DeleteWaiverCustomField WaiverCustomFieldId
        end sub

        public sub Load(WaiverCustomFieldId)
            dim myCon
            set myCon = new WaiverCustomFieldDataConnection

            dim results
            results = myCon.GetWaiverCustomField(WaiverCustomFieldId)

            if ubound(results) > 0 then
                WaiverCustomFieldId = results(WAIVERCUSTOMFIELD_INDEX_ID, 0)
                SiteGuid = results(WAIVERCUSTOMFIELD_INDEX_SITEGUID, 0)
                Name = results(WAIVERCUSTOMFIELD_INDEX_NAME, 0)
                CustomFieldDataType.CustomFieldDataTypeId = results(WAIVERCUSTOMFIELD_INDEX_DATATYPE, 0)
                Notes = results(WAIVERCUSTOMFIELD_INDEX_NOTES, 0)
                Required = results(WAIVERCUSTOMFIELD_INDEX_REQUIRED, 0)
                ValidateDataType = results(WAIVERCUSTOMFIELD_INDEX_VALIDATE, 0)
                DefaultValue = results(WAIVERCUSTOMFIELD_INDEX_DEFAULTVALUE, 0)
                FormLocation = results(WAIVERCUSTOMFIELD_INDEX_FORMLOCATION, 0)
                AdditionalInformation = results(WAIVERCUSTOMFIELD_INDEX_ADDITIONALINFO, 0)

            end if
        end sub
        
        public sub Save()
            dim myCon
            set myCon = new WaiverCustomFieldDataConnection

            myCon.SaveWaiverCustomField WaiverCustomFieldId, Name, CustomFieldDataType.CustomFieldDataTypeId, Notes, Required, ValidateDataType, DefaultValue, FormLocation, AdditionalInformation
        end sub

        ' properties
        public property Get AdditionalInformation
            AdditionalInformation = ms_AdditionalInfo
        end property

        public property Let AdditionalInformation(value)
            ms_AdditionalInfo = value
        end property

        public property Get WaiverCustomFieldId
            WaiverCustomFieldId = mn_Id
        end property

        public property Let WaiverCustomFieldId(value)
            mn_Id = value
        end property

        public property get SiteGuid
            SiteGuid = ms_SiteGuid
        end property

        public property let SiteGuid(value)
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
            DefaultValue = ms_DefaultValue
        end property

        public property Let DefaultValue(value)
            ms_DefaultValue = value
        end property

        public property Get FormLocation
            FormLocation = ms_FormLocation
        end property

        public property Let FormLocation(value)
            ms_FormLocation = value
        end property
    end class
%>