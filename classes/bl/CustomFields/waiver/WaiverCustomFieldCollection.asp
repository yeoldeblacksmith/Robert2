<%
    class WaiverCustomFieldCollection
        ' attributes
        private mo_List

        ' ctor
        public sub Class_Initialize()
            set mo_List = new ArrayList
        end sub

        ' methods
        public sub LoadAll()
            dim myCon
            set myCon = new WaiverCustomFieldDataConnection

            dim fields
            fields = myCon.GetAllActiveWaiverCustomFields()

            if UBound(fields) > 0 then
                for i = 0 to UBound(fields, 2)
                    dim myField
                    set myField = new WaiverCustomField

                    myField.WaiverCustomFieldId = fields(WAIVERCUSTOMFIELD_INDEX_ID, i)
                    myField.SiteGuid = fields(WAIVERCUSTOMFIELD_INDEX_SITEGUID, i)
                    myField.Name = fields(WAIVERCUSTOMFIELD_INDEX_NAME, i)
                    myField.CustomFieldDataType.CustomFieldDataTypeId = fields(WAIVERCUSTOMFIELD_INDEX_DATATYPE, i)
                    myField.Notes = fields(WAIVERCUSTOMFIELD_INDEX_NOTES, i)
                    myField.Required = fields(WAIVERCUSTOMFIELD_INDEX_REQUIRED, i)
                    myField.ValidateDataType = fields(WAIVERCUSTOMFIELD_INDEX_VALIDATE, i)
                    myField.DefaultValue = fields(WAIVERCUSTOMFIELD_INDEX_DEFAULTVALUE, i)
                    myField.FormLocation = fields(WAIVERCUSTOMFIELD_INDEX_FORMLOCATION, i)
                    myField.AdditionalInformation = fields(WAIVERCUSTOMFIELD_INDEX_ADDITIONALINFO, i)
    
                    mo_List.Add myField
                next
            end if
        end sub

        public sub LoadDistinct(includeInActive)
            dim myCon
            set myCon = new WaiverCustomFieldDataConnection

            dim fields
            fields = myCon.GetDistinctCustomFields(includeInActive)

            if UBound(fields) > 0 then
                for i = 0 to UBound(fields, 2)
                    dim myField
                    set myField = new WaiverCustomField

                    myField.WaiverCustomFieldId = fields(WAIVERCUSTOMFIELD_INDEX_ID, i)
                    myField.SiteGuid = fields(WAIVERCUSTOMFIELD_INDEX_SITEGUID, i)
                    myField.Name = fields(WAIVERCUSTOMFIELD_INDEX_NAME, i)
                    myField.CustomFieldDataType.CustomFieldDataTypeId = fields(WAIVERCUSTOMFIELD_INDEX_DATATYPE, i)
                    myField.Notes = fields(WAIVERCUSTOMFIELD_INDEX_NOTES, i)
                    myField.Required = fields(WAIVERCUSTOMFIELD_INDEX_REQUIRED, i)
                    myField.ValidateDataType = fields(WAIVERCUSTOMFIELD_INDEX_VALIDATE, i)
                    myField.DefaultValue = fields(WAIVERCUSTOMFIELD_INDEX_DEFAULTVALUE, i)
                    myField.FormLocation = fields(WAIVERCUSTOMFIELD_INDEX_FORMLOCATION, i)
                    myField.AdditionalInformation = fields(WAIVERCUSTOMFIELD_INDEX_ADDITIONALINFO, i)
    
                    mo_List.Add myField
                next
            end if
        end sub

        ' properties
        public property get Count
            Count = mo_List.Count
        end property

        Public Default Property Get Item(index)
            set Item = mo_List(index)
        end property
    end class
%>