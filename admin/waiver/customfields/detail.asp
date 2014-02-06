<!DOCTYPE html>
<!--#include file="../../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForSiteUsers
%>
<%
    dim field
    set field = new WaiverCustomField

    if request.ServerVariables("REQUEST_METHOD") = "GET" then
        if len(Request.QueryString("id")) > 0 then
            field.WaiverCustomFieldId = request.QueryString("id")
            field.Load request.QueryString("id")
            
            if field.SiteGuid <> MY_SITE_GUID then response.Redirect("../../unauthorized.asp")
        end if
    elseif Request.ServerVariables("REQUEST_METHOD") = "POST" then
        if len(request.Form("txtCustomFieldId")) = 0 then
            field.WaiverCustomFieldId = 0
        else
            field.WaiverCustomFieldId = request.Form("txtCustomFieldId")
        end if
        field.Name = request.Form("txtName")
        field.CustomFieldDataType.CustomFieldDataTypeId = request.Form("ddlType")
        field.Notes = request.Form("txtNotes")
        field.Required = cbool(len(request.Form("chkRequired")) > 0 )
        'field.DefaultValue = request.Form("txtDefaultValue")
        field.FormLocation = request.form("ddlLocation")
        field.AdditionalInformation = request.Form("txtAdditionalInfo")

        field.Save

        response.Redirect "default.asp"
    end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Custom Field Detail</title>
        <link type="text/css" rel="Stylesheet" href="../../../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../../../content/css/admin.css" />
        <link type="text/css" rel="Stylesheet" href="../../../content/css/navmenu.css" />

        <script type="text/javascript" src="../../.../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
        </script>
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Custom Field Detail"

    navmenu.WriteNavigationSection NAVIGATION_NAME_SETTINGS 
%>
        <form method="post">
            <table class="adminTable" style="margin: 0 auto; width: 550px; height: auto;">
                <tr>
                    <td class="adminCell" style=" padding: 10px">
                        <h2 class="center">Custom Field Detail</h2>

                        <input type="hidden" name="txtCustomFieldId" value="<%= field.WaiverCustomFieldId %>" />

                        <table class="tableDefault">
                            <tr>
                                <td>Name:</td>
                                <td>
                                    <input type="text" id="txtName" name="txtName" value="<%= field.Name %>" maxlength="100" class="fullWidth" />
                                </td>
                            </tr>
                            <tr>
                                <td>Type:</td>
                                <td>
                                    <% BuildTypeList field.CustomFieldDataType.CustomFieldDataTypeId %>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">Notes:</td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <% 
                                        Dim editor
	                                    Set editor = New CuteEditor
                                        editor.FilesPath = "../../../../../cute/CuteEditor_Files"
	                                    editor.HelpUrl = "help.asp"
                                        editor.ImageGalleryPath = "../../../../../cute/images/cuteupload"
                                        editor.ConfigurationPath = "../../../../../cute/CuteEditor_Files/Configuration/autoconfigure/simple.config"
                                        editor.Id = "txtNotes"
                                        editor.Text = field.Notes
                                        editor.Draw()
                                    %>
                                </td>
                            </tr>
                            <tr>
                                <td colspan="2">Additional Information:</td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <% 
                                        Dim addInfoEditor
	                                    Set addInfoEditor = New CuteEditor
                                        addInfoEditor.FilesPath = "../../../../../cute/CuteEditor_Files"
	                                    addInfoEditor.HelpUrl = "help.asp"
                                        addInfoEditor.ImageGalleryPath = "../../../../../cute/images/cuteupload"
                                        addInfoEditor.ConfigurationPath = "../../../../../cute/CuteEditor_Files/Configuration/autoconfigure/simple.config"
                                        addInfoEditor.Id = "txtAdditionalInfo"
                                        addInfoEditor.Text = field.AdditionalInformation
                                        addInfoEditor.Draw()
                                    %>
                                </td>
                            </tr>
                            <tr>
                                <td>Required:</td>
                                <td>
                                    <input type="checkbox" name="chkRequired" id="chkRequired" <%= iif(field.Required, "checked=""checked""", "") %> />
                                </td>
                            </tr>
                            <tr>
                                <td>Location:</td>
                                <td>
                                    <% BuildLocationList field.FormLocation%>
                                </td>
                            </tr>
                            <!--<tr>
                                <td>Default:</td>
                                <td>
                                    <input type="text" name="txtDefaultValue" id="txtDefaultValue" value="<%= field.DefaultValue %>" maxlength="1000" class="fullWidth" />
                                </td>
                            </tr>//-->
                            <tr>
                                <td colspan="2" class="center">
                                    <button type="submit">Update</button>&nbsp;
                                    <button type="button" onclick="javascript: window.location.href='default.asp';">Cancel</button>
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
            </table>
        </form>
    </body>
</html>
<% 
    sub BuildTypeList(SelectedValue)
        dim types
        set types = new CustomFieldDataTypeCollection

        types.LoadAll

        with response
            .Write "<select name=""ddlType"" id=""ddlType"">" & vbCrLf
            .Write "<option value=""""></option>" & vbCrLf

            for i = 0 to types.Count - 1
                dim isWritable
                isWritable = true

                if types(i).CustomFieldDataTypeId = FIELDTYPE_YESNO then 
                    if bEdit and SelectedValue = FIELDTYPE_YESNO then
                        ' only allow edits to existing yes/no fields
                    else
                        isWritable = false
                    end if
                end if

                if isWritable then
                    .Write "<option value=""" & types(i).CustomFieldDataTypeId & """ " & iif(types(i).CustomFieldDataTypeId = SelectedValue, "selected=""selected""", "") & ">" & _
                           types(i).Description & "</option>" & vbCrLf
                end if
            next

            .Write "</select>" & vbCrLf
        end with
    end sub

    sub BuildLocationList(SelectedValue)
        with response
            .Write "<select name=""ddlLocation"" id=""ddlLocation"">" & vbCrLf
            .Write "<option value=""""></option>" & vbCrLf

            dim WaiverCustomFields_Location_Description_Index
            WaiverCustomFields_Location_Description_Index = 0

            for each i in WAIVERCUSTOMFIELDS_LOCATION_DESCRIPTION
                .Write "<option value=""" & WaiverCustomFields_Location_Description_Index & """ " & iif(WaiverCustomFields_Location_Description_Index = SelectedValue, "selected=""selected""", "") & ">" & i & "</option>" & vbCrLf
                WaiverCustomFields_Location_Description_Index = WaiverCustomFields_Location_Description_Index + 1
            next

            .Write "</select>" & vbCrLf
        end with

    end sub

%>