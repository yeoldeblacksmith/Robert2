<%
    class CustomDataType
        ' attributes
        dim mn_Id, ms_Description

        ' methods
        public sub Load()
            dim myDLCon
            set myDLCon = new CustomFieldDataTypeDataConnection

            dim results
            results = myDLCon.GetCustomFieldDataType(CustomFieldDataTypeId)

            if ubound(results) > 0 then
                CustomFieldDataTypeId = results(CUSTOMFIELDDATATYPE_INDEX_ID, 0)
                Description = results(CUSTOMFIELDDATATYPE_INDEX_DESC, 0)
            end if
        end sub

        ' properties
        public property Get CustomFieldDataTypeId
            CustomFieldDataTypeId = mn_Id
        end property

        public property Let CustomFieldDataTypeId(value)
            mn_Id = value
        end property

        public property Get Description
            Description = ms_Description
        end property

        public property Let Description(value)
            ms_Description = value
        end property
    end class
%>