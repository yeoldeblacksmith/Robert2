<!DOCTYPE html>
<!--#include file="../../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForSiteUsers
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Waiver Custom Fields</title>
        <link type="text/css" rel="Stylesheet" href="../../../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../../../content/css/admin.css" />
        <link type="text/css" rel="Stylesheet" href="../../../content/css/navmenu.css" />

        <script type="text/javascript" src="../../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
        </script>
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Waiver Custom Fields"

    navmenu.WriteNavigationSection NAVIGATION_NAME_SETTINGS 
%>
        <form method="post">
            <table class="adminTable" style="margin: 0 auto; width: 650px; height: auto;">
                <tr>
                    <td class="adminCell" style=" padding: 10px">
                        <h2 class="center">Custom Fields</h2>
<%
    dim fields
    set fields = new WaiverCustomFieldCollection
    fields.LoadAll

    if fields.Count = 0 then
        response.Write "<p class=""center"">There are no custom fields defined<p/>" & vbCrLf
        response.Write "<a href=""detail.asp"">New Field</a>" & vbCrLf
    else
        BuildFieldTable fields
    end if
%>
                    </td>
                <//tr>
            </table>
        </form>
    </body>
</html>
<%
    sub BuildFieldTable(FieldList)
        with response
            .Write "<table class=""tableDefault fullWidth"">" & vbCrLf
            .Write "<tr>" & vbCrLf
            .Write "<th>Name</th>" & vbCrLf
            .Write "<th>Type</th>" & vbCrLf
            .Write "<th>Req</th>" & vbCrLf
            .Write "<th>Location</th>" & vbCrLf
            '.Write "<th>Default</th>" & vbCrLf
            .Write "<th/>" & vbCrLf
            .Write "</tr>" & vbCrLf

            for i = 0 to FieldList.Count - 1
                .Write "<tr class=""" & iif((i mod 2 = 0), "normal", "alternate") & """>" & vbCrLf
                .Write "<td>" & FieldList(i).Name & "</td>" & vbCrLf
                
                FieldList(i).CustomFieldDataType.Load
                .Write "<td>" & FieldList(i).CustomFieldDataType.Description & "</td>" & vbCrLf

                .Write "<td class=""center""><input type=""checkbox"" readonly=""readonly"" " & iif(FieldList(i).Required, "checked=""checked""", "") & "/></td>" & vbCrLf
                .Write "<td>" & WAIVERCUSTOMFIELDS_LOCATION_DESCRIPTION(FieldList(i).FormLocation) & "</td>" & vbCrLf
                '.Write "<td>" & FieldList(i).DefaultValue & "</td>" & vbCrLf
                .Write "<td>" & vbCrLf
                .Write "<a href=""detail.asp?id=" & FieldList(i).WaiverCustomFieldId & """>Edit</a>&nbsp;|&nbsp;" & vbCrLf
                .Write "<a href=""delete.asp?id=" & FieldList(i).WaiverCustomFieldId & """>Delete</a>" & vbCrLf

                if FieldList(i).CustomFieldDataType.CustomFieldDataTypeId = FIELDTYPE_OPTIONS then
                    .Write "&nbsp;|&nbsp;<a href=""options.asp?id=" & FieldList(i).WaiverCustomFieldId & """>Options</a>" & vbCrLf
                end if
                .Write "</td>" & vbCrLf
                .Write "</tr>" & vbCrLf
            next

            .Write "<tr>" & vbCrLf
            .Write "<td colspan=""3""/>" & vbCrLf
            .Write "<td><a href=""detail.asp"">New Field</a></td>" & vbCrLf
            .Write "</tr>" & vbCrLf

            .Write "</table>" & vbCrLf
        end with
    end sub
%> 