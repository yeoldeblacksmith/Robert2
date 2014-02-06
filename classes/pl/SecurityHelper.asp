<%
    sub AuthorizeForAdminOnly
        ' authenticate/authorize the user
        if Request.Cookies(COOKIE_ROLEID) = AUTHORIZED_ROLE_GLOBALADMIN  then
            ' global admin always have access
        elseif Request.Cookies(COOKIE_ROLEID) = AUTHORIZED_ROLE_ADMIN then
            ' user is authenticated
            if Request.Cookies(COOKIE_SITEID) = MY_SITE_GUID then
                ' user belongs to this site
            else
                'redirect the user to the unauthorized page for their site
                dim correctSite
                set correctSite = new Site
                correctSite.Load Request.Cookies(COOKIE_SITEID)
                response.Redirect(PathCombine(correctSite.VantoraUrl,  "admin/unauthorized.asp"))
            end if
        elseif request.Cookies(COOKIE_ROLEID) = AUTHORIZED_ROLE_MANAGER then
            ' user does not have access
            response.Redirect("../unauthorized.asp")
        else
            dim url
            url = "http://"
            url = url & Request.ServerVariables("HTTP_HOST")
            url = url & Request.ServerVariables("URL")
            if Request.QueryString.Count > 0 then
	            url = url & "?" & Request.QueryString 
            end if

            'Redirect unauthorized users to the logon page.
            Response.Redirect "../default.asp?from=" & Server.URLEncode(url)
        end if
    end sub

    sub AuthorizeForSiteUsers
        ' authenticate/authorize the user
        if Request.Cookies(COOKIE_ROLEID) = AUTHORIZED_ROLE_GLOBALADMIN then
            ' global admin always have access
        elseif Request.Cookies(COOKIE_ROLEID) = AUTHORIZED_ROLE_ADMIN or _
           Request.Cookies(COOKIE_ROLEID) = AUTHORIZED_ROLE_MANAGER then
            ' user is authenticated
            if Request.Cookies(COOKIE_SITEID) = MY_SITE_GUID then
                ' user belongs to this site
            else
                'redirect the user to the unauthorized page for their site
                dim correctSite
                set correctSite = new Site
                correctSite.Load Request.Cookies(COOKIE_SITEID)
                response.Redirect(PathCombine(correctSite.VantoraUrl,  "admin/unauthorized.asp"))
            end if
        else
            dim url
            url = "http://"
            url = url & Request.ServerVariables("HTTP_HOST")
            url = url & Request.ServerVariables("URL")
            if Request.QueryString.Count > 0 then
	            url = url & "?" & Request.QueryString 
            end if

            'Redirect unauthorized users to the logon page.
            Response.Redirect "../default.asp?from=" & Server.URLEncode(url)
        end if
    end sub
%>