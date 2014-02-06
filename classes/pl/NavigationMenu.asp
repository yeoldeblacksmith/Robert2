<%
    class NavigationMenu
        ' attributes
        dim ms_HeadingTitle, mo_Links, mo_Controls, ms_Role

        ' ctor
        public sub Class_Initialize()
            set mo_Links = new NavigationMenuLinkCollection
            set mo_Controls = new ArrayList

            mo_Links.Add BuildHomeGroup()
            if CBool(Settings(SETTING_MODULE_REGISTRATION)) then mo_Links.Add BuildEventGroup()
            mo_Links.Add BuildWaiverGroup()
            mo_Links.Add BuildUserGroup()
            mo_Links.Add BuildSettingGroup()
        end sub

        ' methods
        private function BuildEventGroup()
            dim linkBuilder
            set linkBuilder = new NavigationMenuLink

            linkBuilder.Create NAVIGATION_NAME_EVENT_SCHEDULE, NAVIGATION_URL_EVENT_SCHEDULE, false
            linkBuilder.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_EVENT_NEWREGISTRATION, NAVIGATION_URL_EVENT_NEWREGISTRATION, false)
            linkBuilder.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_EVENT_CALENDAR, NAVIGATION_URL_EVENT_CALENDAR, false)
            'linkBuilder.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_EVENT_AVAILABLEDATES, NAVIGATION_URL_EVENT_AVAILABLEDATES, false)
            linkBuilder.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_EVENT_DAILYREPORT, NAVIGATION_URL_EVENT_DAILYREPORT, false)
            linkBuilder.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_EVENT_PRINT, NAVIGATION_URL_EVENT_PRINT, false)

            set BuildEventGroup = linkBuilder
        end function

        private function BuildHomeGroup()
            dim linkBuilder
            set linkBuilder = new NavigationMenuLink

            set BuildHomeGroup = linkBuilder.CreateNew(NAVIGATION_NAME_HOME, SiteInfo.HomeUrl, false)
        end function

        private function BuildSettingGroup()
            dim linkBuilder
            set linkBuilder = new NavigationMenuLink

            linkBuilder.Create NAVIGATION_NAME_SETTINGS, NAVIGATION_URL_SETTINGS, false

            dim emailLink
            set emailLink = new NavigationMenuLink
            emailLink.Create NAVIGATION_NAME_SETTINGS_EMAIL, NAVIGATION_URL_SETTINGS_EMAIL, false
            emailLink.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_SETTINGS_EMAIL_CANCELLATION, NAVIGATION_URL_SETTINGS_EMAIL_CANCELLATION, false)        
            emailLink.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_SETTINGS_EMAIL_CONFIRMATION, NAVIGATION_URL_SETTINGS_EMAIL_CONFIRMATION, false)        
            'emailLink.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_SETTINGS_EMAIL_INVITE, NAVIGATION_URL_SETTINGS_EMAIL_INVITE, false)        
            emailLink.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_SETTINGS_EMAIL_PAYMENTFAILED, NAVIGATION_URL_SETTINGS_EMAIL_PAYMENTFAILED, false)        
            'emailLink.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_SETTINGS_EMAIL_PAYMENTREMINDER, NAVIGATION_URL_SETTINGS_EMAIL_PAYMENTREMINDER, false)        
            emailLink.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_SETTINGS_EMAIL_REMINDER, NAVIGATION_URL_SETTINGS_EMAIL_REMINDER, false)        
            emailLink.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_SETTINGS_EMAIL_WAIVER, NAVIGATION_URL_SETTINGS_EMAIL_WAIVER, false)        
            emailLink.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_SETTINGS_EMAIL_WAIVERLOOKUP, NAVIGATION_URL_SETTINGS_EMAIL_WAIVERLOOKUP, false)        

            linkBuilder.Children.Add emailLink

            dim regLink
            set regLink = new NavigationMenuLink
            regLink.Create NAVIGATION_NAME_SETTINGS_REGISTRATION, NAVIGATION_URL_SETTINGS_REGISTRATION, false
            regLink.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_SETTINGS_REGISTRATION_AVAILABLEDATES, NAVIGATION_URL_SETTINGS_REGISTRATION_AVAILABLEDATES, false)
            regLink.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_SETTINGS_REGISTRATION_MODIFYSETTINGS, NAVIGATION_URL_SETTINGS_REGISTRATION_MODIFYSETTINGS, false)
            regLink.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_SETTINGS_REGISTRATION_CUSTOMFIELDS, NAVIGATION_URL_SETTINGS_REGISTRATION_CUSTOMFIELDS, false)
            'regLink.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_SETTINGS_REGISTRATION_EVENTTYPES, NAVIGATION_URL_SETTINGS_REGISTRATION_EVENTTYPES, false)

            linkBuilder.Children.Add regLink
            'linkBuilder.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_SETTINGS_INVITES, NAVIGATION_URL_SETTINGS_INVITES, false)

            dim wLink
            set wLink = new NavigationMenuLink
            wLink.Create NAVIGATION_NAME_SETTINGS_WAIVER, NAVIGATION_URL_SETTINGS_WAIVER, false
            wLink.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_SETTINGS_WAIVER_CUSTOMFIELDS, NAVIGATION_URL_SETTINGS_WAIVER_CUSTOMFIELDS, false)
    
            linkBuilder.Children.Add wLink

            set BuildSettingGroup = linkBuilder
        end function

        private function BuildUserGroup()
            dim linkBuilder
            set linkBuilder = new NavigationMenuLink

            linkBuilder.Create NAVIGATION_NAME_USERS, NAVIGATION_URL_USERS, false

            set BuildUserGroup = linkBuilder
        end function

        private function BuildWaiverGroup()
            dim linkBuilder
            set linkBuilder = new NavigationMenuLink

            linkBuilder.Create NAVIGATION_NAME_WAIVERS, NAVIGATION_URL_WAIVERS, false
            linkBuilder.Children.Add linkBuilder.CreateNew(NAVIGATION_NAME_WAIVERS_EXPORT, NAVIGATION_URL_WAIVERS_EXPORT, false)

            set BuildWaiverGroup = linkBuilder
        end function

        public sub WriteNavigationSection(SelectedPage)
            response.Write "<div id=""head_section"">" & vbCrLf
            response.Write "<img id=""logo"" src=""" & LOGO_URL & """ alt=""logo"" style=""width: 180px; height: 90px"" />" & vbCrLf
            response.Write "<div id='login_information'>Welcome, " & Request.Cookies(COOKIE_USER) & " (" & RoleDescription & ") &ndash; <a href='" & SiteInfo.VantoraUrl & "/admin/signout.asp'>Log Off</a> " & "</div>" & vbCrLf
            response.Write "<h2 id=""page_heading"">" & HeadingTitle & "</h2>"
            response.Write "<div id='new_stuff' style='position: absolute; top: 1em; right: 1em'>"
    
            dim fso 
            set fso = Server.CreateObject("Scripting.FileSystemObject")
            dim stream
            set stream = fso.OpenTextFile(PathCombine(Request.ServerVariables("APPL_PHYSICAL_PATH"), "admin\newstuff\newinclude.asp"))
            
            response.Write Replace(stream.ReadAll(), "{{SiteUrl}}", SiteInfo.VantoraUrl)
            stream.Close

            response.Write "</div>"
            response.Write vbCrlf

            response.Write "<ul id=""nav"">" & vbCrLf

            response.Write mo_Links.ToString(false, SelectedPage)

            if FormControlStrings.Count > 0 then
                response.Write "<li style=""float: right"">" & vbCrLf

                dim ControlArray
                ControlArray = FormControlStrings.ToArray()

                'for i = 0 to FormControlStrings.Count - 1
                for i = 0 to UBound(ControlArray)
                    'response.Write FormControlStrings(i) & vbCrLf
                    response.Write ControlArray(i) & vbCrLf
                next 

                response.Write "</li>" & vbCrLf
            end if

            response.Write "</ul>" & vbCrLf
            response.Write "</div>" & vbCrLf
        end sub

        ' properties
        public property get HeadingTitle
            HeadingTitle = ms_HeadingTitle
        end property

        public property get RoleDescription
            Dim myRole
            set myRole = new AuthorizationRole
            myRole.RoleId = Request.Cookies(COOKIE_ROLEID)
            myRole.Load
            RoleDescription = myRole.Description
        end property

        public property let HeadingTitle(value)
            ms_HeadingTitle = value
        end property

        public property get FormControlStrings
            set FormControlStrings = mo_Controls
        end property

        public property let FormControlStrings(value)
            set mo_Controls = value
        end property

    end class
%>