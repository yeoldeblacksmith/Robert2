<!DOCTYPE html>
<!--#include file="../../../classes/IncludeList.asp" -->
<%
    dim field, bEdit
    set field = new RegistrationCustomField
    bEdit = false

    if request.ServerVariables("REQUEST_METHOD") = "GET" then
        if len(Request.QueryString("id")) > 0 then
            bEdit = true
            field.RegistrationCustomFieldId = request.QueryString("id")
            field.Load
            
            if field.SiteGuid <> MY_SITE_GUID then response.Redirect("../../unauthorized.asp")
        end if
    elseif Request.ServerVariables("REQUEST_METHOD") = "POST" then
        if len(request.Form("txtCustomFieldId")) = 0 then
            field.RegistrationCustomFieldId = 0
        else
            field.RegistrationCustomFieldId = request.Form("txtCustomFieldId")
        end if
        field.Name = request.Form("txtName")
        field.CustomFieldDataType.CustomFieldDataTypeId = request.Form("ddlType")
        field.Notes = Trim(request.Form("txtNotes"))
        field.Required = cbool(len(request.Form("chkRequired")) > 0 )
        'field.DefaultValue = request.Form("txtDefaultValue")
        field.AdditionalInformation = request.Form("txtAdditionalInfo")
        field.Sequence = request.Form("txtSequence")
        field.PaymentTypeId = request.Form(SETTING_REGISTRATION_PAYMENTTYPE)

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

        <script type="text/javascript" src="../../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                togglePaymentTypeRowVisibility();

                $("#ddlType").on("change", function () { togglePaymentTypeRowVisibility() });
            });

            function validateForm() {
                var valid = !!(validateRequired("txtName", "valNameRequired") & validateRequired("ddlType", "valTypeRequired"));
                return valid;
            }

            function validateRequired(fieldControlId, validatorControlId) {
                var valid = true;

                if ($("#" + fieldControlId).val().length == 0) {
                    var valid = false;
                    $("#" + validatorControlId).css("display", "block");
                } else {
                    $("#" + validatorControlId).css("display", "none");
                }

                return valid;
            }

            function togglePaymentTypeRowVisibility() {
                if ($("#ddlType").val() == "<%= FIELDTYPE_OPTIONS %>") {
                    $("#rowPaymentType").css("display", "table-row");
                } else {
                    $("#rowPaymentType").css("display", "none");
                    $("PaymentType").val("<%= PAYMENTTYPE_NONE %>");
                }
            }
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

                        <input type="hidden" name="txtCustomFieldId" value="<%= field.RegistrationCustomFieldId %>" />
                        <input type="hidden" name="txtSequence" value="<%= field.Sequence %>" />

                        <table class="tableDefault">
                            <tr>
                                <td>Name: <span style="color: red">*</span></td>
                                <td>
                                    <input type="text" id="txtName" name="txtName" value="<%= field.Name %>" maxlength="100" class="fullWidth" /><br />
                                    <span id="valNameRequired" style="display: none; color: red">Required</span>
                                </td>
                            </tr>
                            <tr>
                                <td>Type: <span style="color: red">*</span></td>
                                <td>
                                    <% BuildTypeList field.CustomFieldDataType.CustomFieldDataTypeId %>
                                    <br />
                                    <span id="valTypeRequired" style="display: none; color: red">Required</span>
                                </td>
                            </tr>
                            <tr id="rowPaymentType">
                                <td>Payment Type:</td>
                                <td>
                                    <% BuildPaymentTypeList field.PaymentTypeId %>
                                </td>
                            </tr>
                            <tr>
                                <td>Required:</td>
                                <td>
                                    <input type="checkbox" name="chkRequired" id="chkRequired" <%= iif(field.Required, "checked=""checked""", "") %> />
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
                                        editor.ImageGalleryPath = "../../../content/images/cuteupload"
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
                                        addInfoEditor.ImageGalleryPath = "../../../content/images/cuteupload"
                                        addInfoEditor.ConfigurationPath = "../../../../../cute/CuteEditor_Files/Configuration/autoconfigure/simple.config"
                                        addInfoEditor.Id = "txtAdditionalInfo"
                                        addInfoEditor.Text = field.AdditionalInformation
                                        addInfoEditor.Draw()
                                    %>
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
                                    <button type="submit" onclick="return validateForm()">Update</button>&nbsp;
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
%>