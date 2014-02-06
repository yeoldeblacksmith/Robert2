<%
    class NavigationMenuLink
        'attributes
        dim ms_Name, ms_Url, mb_Selected, _
            mo_Children

        ' ctor
        public sub Class_Initialize()
            ms_Name = ""
            ms_Url = ""
            mb_Selected = false
        
            set mo_Children = new NavigationMenuLinkCollection
        end sub

        ' method
        public sub Create(sName, sURL, bSelected)
            Name = sName
            URL = sURL
            Selected = bSelected
        end sub

        public function CreateNew(sName, sURL, bSelected)
            dim link
            set link = new NavigationMenuLink

            link.Create sName, sURL, bSelected

            set CreateNew = link
        end function

        public function ToString(SelectedPage)
            dim output

            output = "<li" & iif(Selected, " class=""current""", "") & ">" & vbCrLf 
            
            if left(lcase(URL), 11) = "javascript:" or left(lcase(URL), 4) = "http" then
                output = output & "<a href=""" & URL & """>" & Name & "</a>" & vbCrLf
            else
                output = output & "<a href=""" & PathCombine(SiteInfo.VantoraUrl, URL) & """>" & Name & "</a>" & vbCrLf
            end if

            if Children.Count > 0 then
                output = output & Children.ToString(true, SelectedPage)
            end if

            output = output & "</li>" & vbCrLf

            ToString = output
        end function

        public property Get Name
            Name = ms_Name
        end property

        public property Let Name(value)
            ms_Name = value
        end property

        public property Get URL
            URL = ms_Url
        end property

        public property Let URL(value)  
            ms_Url = value
        end property

        public property Get Selected
            Selected = mb_Selected
        end property

        public property let Selected(value)
            mb_Selected = value
        end property

        public property get Children
            set Children = mo_Children
        end property

        public property let Children(value)
            set mo_Children = value
        end property
    end class
%>