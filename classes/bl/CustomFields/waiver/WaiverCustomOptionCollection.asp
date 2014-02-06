<%
    class WaiverCustomOptionCollection
        ' attributes
        private mo_List

        ' ctor
        public sub Class_Initialize()
            set mo_List = new ArrayList
        end sub

        ' methods
        public sub LoadByFieldId(FieldId)
            dim myCon
            set myCon = new WaiverCustomOptionDataConnection

            dim options
            options = myCon.GetAllActiveWaiverCustomOptionsByFieldId(FieldId)

            if UBound(options) > 0 then
                for i = 0 to UBound(options, 2)
                    dim myOption
                    set myOption = new WaiverCustomOption

                    myOption.WaiverCustomFieldId = options(WAIVERCUSTOMOPTION_INDEX_ID, i)
                    myOption.Sequence = options(WAIVERCUSTOMOPTION_INDEX_SEQUENCE, i)
                    myOption.Text = options(WAIVERCUSTOMOPTION_INDEX_TEXT, i)
                    myOption.Value = options(WAIVERCUSTOMOPTION_INDEX_VALUE, i)

                    mo_List.Add myOption
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