<%
    class RegistrationCustomFieldCollection
        ' attributes
        private mo_List

        ' ctor
        public sub Class_Initialize()
            set mo_List = new ArrayList
        end sub

        ' methods
        public sub LoadAll()
            dim myCon
            set myCon = new RegistrationCustomFieldDataConnection

            dim fields
            fields = myCon.GetAllActiveRegistrationCustomFields()

            if UBound(fields) > 0 then
                for i = 0 to UBound(fields, 2)
                    dim myField
                    set myField = new RegistrationCustomField

                    myField.RegistrationCustomFieldId = fields(REGISTRATIONCUSTOMFIELD_INDEX_ID, i)
                    myField.SiteGuid = fields(REGISTRATIONCUSTOMFIELD_INDEX_SITEGUID, i)
                    myField.Name = fields(REGISTRATIONCUSTOMFIELD_INDEX_NAME, i)
                    myField.CustomFieldDataType.CustomFieldDataTypeId = fields(REGISTRATIONCUSTOMFIELD_INDEX_DATATYPE, i)
                    myField.Notes = fields(REGISTRATIONCUSTOMFIELD_INDEX_NOTES, i)
                    myField.Required = fields(REGISTRATIONCUSTOMFIELD_INDEX_REQUIRED, i)
                    myField.ValidateDataType = fields(REGISTRATIONCUSTOMFIELD_INDEX_VALIDATE, i)
                    myField.DefaultValue = fields(REGISTRATIONCUSTOMFIELD_INDEX_DEFAULTVALUE, i)
                    myField.AdditionalInformation = fields(REGISTRATIONCUSTOMFIELD_INDEX_ADDITIONALINFO, i)
                    myField.Sequence = fields(REGISTRATIONCUSTOMFIELD_INDEX_SEQUENCE, i)
                    myField.PaymentTypeId = fields(REGISTRATIONCUSTOMFIELD_INDEX_PAYMENTTYPE, i)

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