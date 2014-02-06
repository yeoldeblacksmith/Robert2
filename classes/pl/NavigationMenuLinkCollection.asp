<%
    class NavigationMenuLinkCollection
        'attributes
        private mo_List

        ' ctor
        public sub Class_Initialize()
            set mo_List = new ArrayList
        end sub

        'methods
        public sub Add(Link)
            mo_List.Add link
        end sub

        public sub Remove(Link)
            mo_List.Remove link
        end sub 

        public function ToString(WriteContainerTag, SelectedPageName)
            dim output

            if WriteContainerTag then output = "<ul>" & vbCrlf

            for i = 0 to Count - 1
                if Item(i).Name = SelectedPageName then 
                    Item(i).Selected = true
                end if

                output = output & Item(i).ToString(SelectedPage)
            next

            if WriteContainerTag then output = output & "</ul>" & vbCrLf

            ToString = output
        end function

        'properties
        public property get Count
            Count = mo_List.Count
        end property

        Public Default Property Get Item(index)
            set Item = mo_List(index)
        end property
    end class
%>