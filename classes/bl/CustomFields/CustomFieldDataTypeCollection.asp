<%
    class CustomFieldDataTypeCollection
        ' attributes
        private mo_List

        ' ctor
        public sub Class_Initialize()
            set mo_List = new ArrayList
        end sub

        ' methods
        public sub LoadAll()
            dim myCon
            set myCon = new CustomFieldDataTypeDataConnection

            dim types
            types = myCon.GetAllCustomFieldDataTypes()

            if UBound(types) > 0 then
                for i = 0 to UBound(types, 2)
                    dim myType
                    set myType = new CustomDataType

                    myType.CustomFieldDataTypeId = types(CUSTOMFIELDDATATYPE_INDEX_ID, i)
                    myType.Description = types(CUSTOMFIELDDATATYPE_INDEX_DESC, i)

                    mo_List.Add myType
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