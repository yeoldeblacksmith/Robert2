<%
    class WaiverCustomValueCollection
        ' attributes
        private mo_List

        ' ctor
        public sub Class_Initialize()
            set mo_List = new ArrayList
        end sub

        ' methods
        public sub DeleteByWaiverId(WaiverId)
            dim myCon
            set myCon = new WaiverCustomValueDataConnection

            myCon.DeleteWaiverCustomValueByWaiverId WaiverId
        end sub

        public sub LoadByWaiverId(WaiverId)
            dim myCon
            set myCon = new WaiverCustomValueDataConnection

            dim values
            values = myCon.GetAllWaiverCustomValuesByWaiverId(WaiverId)

            if UBound(values) > 0 then
                for i = 0 to UBound(values, 2)
                    dim myValue
                    set myValue = new WaiverCustomValue

                    myValue.WaiverId = values(WAIVERCUSTOMVALUE_INDEX_WAIVERID, i)
                    myValue.WaiverCustomFieldId = values(WAIVERCUSTOMVALUE_INDEX_FIELDID, i)
                    myValue.Value = values(WAIVERCUSTOMVALUE_INDEX_VALUE, i)

                    mo_List.Add myValue
                next
            end if
        end sub

        public sub LoadActiveByWaiverId(WaiverId)
            dim myCon
            set myCon = new WaiverCustomValueDataConnection

            dim values
            values = myCon.GetAllActiveWaiverCustomValuesByWaiverId(WaiverId)

            if UBound(values) > 0 then
                for i = 0 to UBound(values, 2)
                    dim myValue
                    set myValue = new WaiverCustomValue

                    myValue.WaiverId = values(WAIVERCUSTOMVALUE_INDEX_WAIVERID, i)
                    myValue.WaiverCustomFieldId = values(WAIVERCUSTOMVALUE_INDEX_FIELDID, i)
                    myValue.Value = values(WAIVERCUSTOMVALUE_INDEX_VALUE, i)

                    mo_List.Add myValue
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