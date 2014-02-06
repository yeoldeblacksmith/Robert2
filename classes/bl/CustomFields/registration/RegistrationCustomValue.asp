<%
    class RegistrationCustomValue
        ' attributes
        dim mn_EventId, mn_FieldId, ms_Value

        ' methods
        public sub Load(nEventId, nRegistrationCustomFieldId)
            dim myCon
            set myCon = new RegistrationCustomValueDataConnection

            dim results
            results = myCon.GetRegistrationCustomValue(nEventId, nRegistrationCustomFieldId)

            if UBound(results) > 0 then
                EventId = results(REGISTRATIONCUSTOMVALUE_INDEX_EVENTID, 0)
                RegistraionCustomFieldId = results(REGISTRATIONCUSTOMVALUE_INDEX_FIELDID, 0)
                Value = results(REGISTRATIONCUSTOMVALUE_INDEX_VALUE, 0)
            end if
        end sub

        public sub Save()
            dim myCon
            set myCon = new RegistrationCustomValueDataConnection

            myCon.SaveRegistrationCustomValue EventId, RegistrationCustomFieldId, Value
        end sub

        ' properties
        public property get EventId
            EventId = mn_EventId
        end property

        public property let EventId(value)  
            mn_EventId = value
        end property

        public property get RegistrationCustomFieldId
            RegistrationCustomFieldId = mn_FieldId
        end property

        public property let RegistrationCustomFieldId(value)
            mn_FieldId = value
        end property

        public property get Value
            Value = ms_Value
        end property

        public property let Value(val)
            ms_Value = val
        end property

    end class
%>