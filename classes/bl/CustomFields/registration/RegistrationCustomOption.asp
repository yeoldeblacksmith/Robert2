<%
    class RegistrationCustomOption
        ' attributes
        dim mn_FieldId, mn_Sequence, ms_Text, _
            ms_Value

        'ctor
        public sub Class_Initialize
            ms_Value = 0
        end sub

        ' methods
        public sub Delete()
            dim myCon
            set myCon = new RegistrationCustomOptionDataConnection

            myCon.DeleteRegistrationCustomOption RegistrationCustomFieldId, Sequence
        end sub

        public sub Load()
            dim myCon
            set myCon = new RegistrationCustomOptionDataConnection

            dim results
            results = myCon.GetRegistrationCustomOption(RegistrationCustomFieldId, Sequence)

            if ubound(results) > 0 then
                RegistrationCustomFieldId = results(REGISTRATIONCUSTOMOPTION_INDEX_ID, 0)
                Sequence = results(REGISTRATIONCUSTOMOPTION_SEQUENCE, 0)
                Text = results(REGISTRATIONCUSTOMOPTION_TEXT, 0)
                Value = results(REGISTRATIONCUSTOMOPTION_VALUE, 0)
            end if
        end sub

        public sub Save()
            dim myCon
            set myCon = new RegistrationCustomOptionDataConnection

            if len(Sequence) = 0 then Sequence = myCon.GetNextRegistrationCustomOptionSequence(RegistrationCustomFieldId)
            myCon.SaveRegistrationCustomOption RegistrationCustomFieldId, Sequence, Text, Value
        end sub

        ' properties
        public property get RegistrationCustomFieldId
            RegistrationCustomFieldId = mn_FieldId
        end property

        public property let RegistrationCustomFieldId(value)
            mn_FieldId = value
        end property

        public property get Sequence
            Sequence = mn_Sequence
        end property

        public property let Sequence(value)
            mn_Sequence = value
        end property

        public property get Text
            Text = ms_Text
        end property

        public property let Text(value)
            ms_Text = value
        end property

        public property get Value
            if isnull(ms_Value) or len(ms_value) = 0 then
                Value = 0
            else
                Value = FormatNumber(cdbl(ms_Value), 2)
            end if
        end property

        public property let Value(val)
            ms_Value = val
        end property                
    end class
%>