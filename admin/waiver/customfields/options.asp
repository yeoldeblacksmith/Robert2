<!DOCTYPE html>
<!--#include file="../../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForSiteUsers
%>
<%
    dim fieldId, editSequence
    
    if Request.ServerVariables("REQUEST_METHOD") = "GET" then
        fieldId = request.QueryString("id")
        editSequence = ""
    else
        fieldId = request.Form("txtCustomFieldId")
        editSequence = request.Form("txtSelected")

        dim myOption
        set myOption = new WaiverCustomOption
        myOption.WaiverCustomFieldId = fieldId
        myOption.Text = request.Form("txtText")
        'myOption.Value = request.Form("txtValue")
        myOption.Value = ""

        select case request.Form("txtAction")
            case "add"
                myOption.Save
            case "cancel"
                editSequence = ""
            case "delete"
                myOption.Sequence = editSequence
                myOption.Delete

                editSequence = ""
            case "edit"
                ' do nothing
            case "update"
                myOption.Sequence = editSequence
                myOption.Save

                editSequence = ""
        end select
    end if    
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Custom Field Options</title>
        <link type="text/css" rel="Stylesheet" href="../../../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../../../content/css/admin.css" />
        <link type="text/css" rel="Stylesheet" href="../../../content/css/navmenu.css" />

        <script type="text/javascript" src="../../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
            function add() {
                $("#txtAction").val("add");
                document.forms[0].submit();
            }

            function cancel() {
                $("#txtAction").val("cancel");
                document.forms[0].submit();
            }

            function del(sequence) {
                if (confirm("Are you sure?")) {
                    $("#txtAction").val("delete");
                    $("#txtSelected").val(sequence);
                    document.forms[0].submit();
                }
            }
            
            function edit(sequence) {
                $("#txtAction").val("edit");
                $("#txtSelected").val(sequence);
                document.forms[0].submit();
            }

            function update(sequence) {
                $("#txtAction").val("update");
                $("#txtSelected").val(sequence);
                document.forms[0].submit();
            }

        </script>
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Delete Custom Field"

    navmenu.WriteNavigationSection NAVIGATION_NAME_SETTINGS 
%>
        <form method="post">
            <table class="adminTable" style="margin: 0 auto; width: 550px; height: auto;">
                <tr>
                    <td class="adminCell" style=" padding: 10px">
                        <h2 class="center">Field Options</h2>
                        <input type="hidden" name="txtCustomFieldId" value="<%= fieldId %>" />
                        <input type="hidden" name="txtSelected" id="txtSelected" />
                        <input type="hidden" name="txtAction" id="txtAction" />

                        <table class="fullWidth tableDefault">
                            <tr>
                                <th>Text</th>
                                <!--<th>Value</th>//-->
                                <th />
                            </tr>
                            <% BuildOptionGrid %>
                        </table>

                        <br /><br />
                        <a href="default.asp">Return to fields</a>
                    </td>
                </tr>
            </table>
        </form>
    </body>
</html>
<% 
    sub BuildOptionGrid()
        dim options
        set options = new WaiverCustomOptionCollection

        options.LoadByFieldId fieldId

        with Response
            for i = 0 to options.Count - 1
                dim selected
                if len(editSequence) = 0 then
                    selected = false
                else
                    selected = cbool(options(i).Sequence = cint(editSequence))
                end if

                .Write "<tr class=""" & iif(cbool(i mod 2 = 0), "normal", "alternate") & """>" & vbCrLf

                ' text cell
                .Write "<td>" & vbCrLf
                if selected then
                    .Write "<input type=""text"" name=""txtText"" id=""txtText"" class=""fullWidth"" maxlength=""100"" value=""" & options(i).Text & """/>" & vbCrLf
                else
                    .Write options(i).Text
                end if
                .Write "</td>" & vbCrLf

                ' value cell
                '.Write "<td>" & vbCrLf
                'if selected then
                '    .Write "<input type=""text"" name=""txtValue"" id=""txtValue"" class=""fullWidth"" maxlength=""100"" value=""" & options(i).Value & """/>" & vbCrLf
                'else
                '    .Write options(i).Value
                'end if
                '.Write "</td>" & vbCrLf

                ' action cell
                .Write "<td>" & vbCrLf
                if selected then
                    .Write "<a href=""javascript: update(" & options(i).Sequence & ");"">Update</a>&nbsp;|&nbsp;" & vbCrLf
                    .Write "<a href=""javascript: cancel();"">Cancel</a>" & vbCrLf
                else
                    .Write "<a href=""javascript: edit(" & options(i).Sequence & ");"">Edit</a>&nbsp;|&nbsp;" & vbCrLf
                    .Write "<a href=""javascript: del(" & options(i).Sequence & ");"">Delete</a>" & vbCrLf
                end if                
                .Write "</td>" & vbCrLf

                .Write "</tr>" & vbCrLf
            next

            ' show the footer row
            if len(editSequence) = 0 then
                .Write "<tr>" & vbCrLf
                .Write "<td>" & vbCrLf
                .Write "<input type=""text"" name=""txtText"" id=""txtText"" class=""fullWidth"" maxlength=""100"" value=""""/>" & vbCrLf
                .Write "</td>" & vbCrLf
                '.Write "<td>" & vbCrLf
                '.Write "<input type=""text"" name=""txtValue"" id=""txtValue"" class=""fullWidth"" maxlength=""100"" value=""""/>" & vbCrLf
                '.Write "</td>" & vbCrLf
                .Write "<td>" & vbCrLf
                .Write "<a href=""javascript: add();"">Add</a>" & vbCrLf
                .Write "</td>" & vbCrLf
                .Write "</tr>" & vbCrLf
            end if
        end with
    end sub
%>