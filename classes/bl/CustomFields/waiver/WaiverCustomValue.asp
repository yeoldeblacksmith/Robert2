<%
    class WaiverCustomValue
        ' attributes
        dim mn_WaiverId, mn_FieldId, ms_Value

        ' methods
        public sub Save()
            dim myCon
            set myCon = new WaiverCustomValueDataConnection

            myCon.SaveWaiverCustomValue WaiverId, WaiverCustomFieldId, Value
        end sub

        ' properties
        public property get CustomFieldDataType
            dim myWaiverCustomField    
            set myWaiverCustomField = new WaiverCustomField
            myWaiverCustomField.Load WaiverCustomFieldId
            CustomFieldDataType = myWaiverCustomField.CustomFieldDataType.CustomFieldDataTypeId
        end property

        public property get Name
            dim myWaiverCustomField    
            set myWaiverCustomField = new WaiverCustomField
            myWaiverCustomField.Load WaiverCustomFieldId
            Name = myWaiverCustomField.Name
        end property

        public property get WaiverId
            WaiverId = mn_WaiverId
        end property

        public property let WaiverId(value)  
            mn_WaiverId = value
        end property

        public property get WaiverCustomFieldId
            WaiverCustomFieldId = mn_FieldId
        end property

        public property let WaiverCustomFieldId(value)
            mn_FieldId = value
        end property

        public property get LocationInWaiver
            dim myWaiverCustomField    
            set myWaiverCustomField = new WaiverCustomField
            myWaiverCustomField.Load WaiverCustomFieldId
            LocationInWaiver = myWaiverCustomField.FormLocation
        end property

        public property get Value
            Value = ms_Value
        end property

        public property let Value(val)
            ms_Value = val
        end property
    end class
%>