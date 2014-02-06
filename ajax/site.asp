<!--#include file="../classes/includelist.asp"-->
<%
    CONST ACTION_GENERATEGUID = "0"

    select case request.QueryString(QUERYSTRING_VAR_ACTION)
        case ACTION_GENERATEGUID
            WriteNewGuid
    end select

    private sub WriteNewGuid()
        dim typeLib
        set typeLib = server.CreateObject("Scriptlet.TypeLib")

        response.Write left(cstr(typeLib.Guid), 38)

        set typeLib = nothing
    end sub
%>