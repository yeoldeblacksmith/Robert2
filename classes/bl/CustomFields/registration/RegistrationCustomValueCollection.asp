<%
    class RegistrationCustomValueCollection
        ' attributes
        private mo_List

        ' ctor
        public sub Class_Initialize()
            set mo_List = new ArrayList
        end sub

        ' methods
        public sub Add(oValue)
            mo_List.Add oValue
        end sub

        public sub LoadByEventId(EventId)
            dim myCon
            set myCon = new RegistrationCustomValueDataConnection

            dim values
            values = myCon.GetAllRegistrationCustomValuesByEventId(EventId)

            if UBound(values) > 0 then
                for i = 0 to UBound(values, 2)
                    dim myValue
                    set myValue = new RegistrationCustomValue

                    myValue.EventId = values(REGISTRATIONCUSTOMVALUE_INDEX_EVENTID, i)
                    myValue.RegistrationCustomFieldId = values(REGISTRATIONCUSTOMVALUE_INDEX_FIELDID, i)
                    myValue.Value = values(REGISTRATIONCUSTOMVALUE_INDEX_VALUE, i)

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