<%
    class RegistrationCustomOptionCollection
        ' attributes
        private mo_List

        ' ctor
        public sub Class_Initialize()
            set mo_List = new ArrayList
        end sub

        ' methods
        public sub LoadByFieldId(FieldId)
            dim myCon
            set myCon = new RegistrationCustomOptionDataConnection

            dim options
            options = myCon.GetAllActiveRegistrationCustomOptionsByFieldId(FieldId)

            if UBound(options) > 0 then
                for i = 0 to UBound(options, 2)
                    dim myOption
                    set myOption = new RegistrationCustomOption

                    myOption.RegistrationCustomFieldId = options(REGISTRATIONCUSTOMOPTION_INDEX_ID, i)
                    myOption.Sequence = options(REGISTRATIONCUSTOMOPTION_SEQUENCE, i)
                    myOption.Text = options(REGISTRATIONCUSTOMOPTION_TEXT, i)
                    myOption.Value = options(REGISTRATIONCUSTOMOPTION_VALUE, i)

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