<%
    class WaiverCustomOption
        ' attributes
        dim mn_FieldId, mn_Sequence, ms_Text, _
            ms_Value

        ' methods
        public sub Delete()
            dim myCon
            set myCon = new WaiverCustomOptionDataConnection

            myCon.DeleteWaiverCustomOption WaiverCustomFieldId, Sequence
        end sub

        public sub Load()
            dim myCon
            set myCon = new WaiverCustomOptionDataConnection

            dim results
            results = myCon.GetWaiverCustomOption(WaiverCustomFieldId, Sequence)

            if ubound(results) > 0 then
                WaiverCustomFieldId = results(WAIVERCUSTOMOPTION_INDEX_ID, 0)
                Sequence = results(WAIVERCUSTOMOPTION_INDEX_SEQUENCE, 0)
                Text = results(WAIVERCUSTOMOPTION_INDEX_TEXT, 0)
                Value = results(WAIVERCUSTOMOPTION_INDEX_VALUE, 0)
            end if
        end sub

        public sub Save()
            dim myCon
            set myCon = new WaiverCustomOptionDataConnection

            if len(Sequence) = 0 then Sequence = myCon.GetNextWaiverCustomOptionSequence(WaiverCustomFieldId)
            myCon.SaveWaiverCustomOption WaiverCustomFieldId, Sequence, Text, Value
        end sub

        ' properties
        public property get WaiverCustomFieldId
            WaiverCustomFieldId = mn_FieldId
        end property

        public property let WaiverCustomFieldId(value)
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
            Value = ms_Value
        end property

        public property let Value(val)
            ms_Value = val
        end property
    end class
%>