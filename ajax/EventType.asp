<!--#include file="../classes/IncludeList.asp" -->
<%
    const EVENTTYPEAJAX_ACTION_GETTYPETABLE = "0"
    const EVENTTYPEAJAX_ACTION_SAVETYPE = "1"
    const EVENTTYPEAJAX_ACTION_DELETETYPE = "2"

    select case Request.QueryString(QUERYSTRING_VAR_ACTION)
        case EVENTTYPEAJAX_ACTION_GETTYPETABLE
            BuildEventTypeTable
        case EVENTTYPEAJAX_ACTION_SAVETYPE
            SaveType Request.QueryString(QUERYSTRING_VAR_DESCRIPTION)
            BuildEventTypeTable
        case EVENTTYPEAJAX_ACTION_DELETETYPE
            DeleteType Request.QueryString(QUERYSTRING_VAR_ID)
            BuildEventTypeTable
    end select

    '********************************************************************
    ' load the table with all active event types
    '********************************************************************
    sub BuildEventTypeTable()
        Response.Write "<table cellpadding=""3"" cellspacing=""2"" border=""1"" style=""margin: 0 auto"">"
        Response.Write "<tr>"
        Response.Write "<td class=""heading"">Description</td>"
        Response.Write "<td class=""heading""/>"
        Response.Write "</tr>"

        dim types
        set types = new EventTypeCollection

        types.LoadAllActive
        
        for i = 0 to types.Count - 1
            dim myType
            set myType = types(i)

            response.Write "<tr>"

            response.Write "<td>"
            response.Write myType.Description
            response.Write "</td>"

            response.Write "<td>"
            response.Write "<a href='#' onclick='deleteType(" & myType.EventTypeId & ")'>Delete</a>"
            response.Write "</td>"

            response.Write "</tr>"
        next

        Response.Write "<tr>"
        Response.Write "<td>"
        Response.Write "<input type=""text"" maxlength=""50"" name=""Description"" style=""width: 98%"" />"
        Response.Write "</td>"
        Response.Write "<td>"
        Response.Write "<a href='#' onclick='saveType()'>Add</a>"
        Response.Write "</td>"
        Response.Write "</tr>"
        Response.Write "</table>"

    end sub

    '********************************************************************
    ' delete the specified event type
    '********************************************************************
    sub DeleteType(id)
        dim myType
        set myType = new EventType

        myType.EventTypeId = id
        myType.Delete
    end sub

    '********************************************************************
    ' create a new event type
    '********************************************************************
    sub SaveType(description)
        dim myType
        set myType = new EventType

        myType.EventTypeId = 0
        myType.Description = description
        myType.Save
    end sub
%>