<!DOCTYPE html>
<!--#include file="../../../classes/IncludeList.asp" -->
<%
    dim fieldId, editSequence, showValueColumn
    showValueColumn = false

    if Request.ServerVariables("REQUEST_METHOD") = "GET" then
        fieldId = request.QueryString("id")
        editSequence = ""
    else
        fieldId = request.Form("txtCustomFieldId")
        editSequence = request.Form("txtSelected")

        dim myOption
        set myOption = new RegistrationCustomOption
        myOption.RegistrationCustomFieldId = fieldId
        myOption.Text = request.Form("txtText")
        myOption.Value = request.Form("txtValue")
        'myOption.Value = ""

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

    dim myField : set myField = new RegistrationCustomField
    myField.LoadById fieldId

    showValueColumn = cbool(cstr(myField.PaymentTypeId) > PAYMENTTYPE_NONE)
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Custom Field Options</title>
        <link type="text/css" rel="Stylesheet" href="../../../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../../../content/css/admin.css" />
        <link type="text/css" rel="Stylesheet" href="../../../content/css/navmenu.css" />
        <link type="text/css" rel="stylesheet" href="../../../content/css/jquery.ui.core.css" />

        <script type="text/javascript" src="../../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="../../../content/js/jquery.ui.core.min.js"></script>
        <script type="text/javascript" src="../../../content/js/jquery.ui.widget.min.js"></script>
        <script type="text/javascript" src="../../../content/js/jquery.ui.mouse.min.js"></script>
        <script type="text/javascript" src="../../../content/js/jquery.ui.sortable.min.js"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                $(".sortable").sortable({ handle: '.dragHandle' });

                $(".sortable").on("sortupdate", function (event, ui) {
                    var rows = $(".sortable > tr");
                    for (i = 0; i < rows.length; i++) {
                        rows.eq(i).removeClass("normal");
                        rows.eq(i).removeClass("alternate");

                        if (i % 2 == 0) {
                            rows.eq(i).addClass("normal");
                        } else {
                            rows.eq(i).addClass("alternate");
                        }

                        //$.ajax({
                        //    url: '../../../ajax/RegistrationCustomField.asp?act=0&id=' + rows.eq(i).data("fieldid") + '&s=' + i,
                        //    error: function () { alert("problems encounered"); },
                        //    async: false
                        //});
                    }
                });
            });

            function add() {
                if (validateValue()) {
                    $("#txtAction").val("add");
                    document.forms[0].submit();
                }
            }

            function cancel() {
                $("#txtAction").val("cancel");
                document.forms[0].submit();
            }

            function del(sequence) {
                if(confirm("Are you sure?")){
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
                if (validateValue()) {
                    $("#txtAction").val("update");
                    $("#txtSelected").val(sequence);
                    document.forms[0].submit();
                }
            }

            function validateValue() {
                var valid = true;
                var valueField = $("#txtValue");

                $("#valValueFormat").css("display", "none");

                if (valueField.length == 1) {
                    if (valueField.val().length > 0) {
                        if(isNaN(valueField.val())) {
                            valid = false;
                            $("#valValueFormat").css("display", "block");
                        }
                    } else {
                        valueField.val("0");
                    }
                }

                return valid;
            }
        </script>
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Custom Field Options"

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
                                <th />
                                <th>Text</th>
                                <th<%= iif(showValueColumn, "", " style=""display: none""") %>>Fee</th>
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
        set options = new RegistrationCustomOptionCollection

        options.LoadByFieldId fieldId

        with Response
            .Write "<tbody class=""sortable"">" & vbCrLf

            for i = 0 to options.Count - 1
                dim selected
                if len(editSequence) = 0 then
                    selected = false
                else
                    selected = cbool(options(i).Sequence = cint(editSequence))
                end if

                .Write "<tr class=""" & iif(cbool(i mod 2 = 0), "normal", "alternate") & """ data-fieldid=""" & fieldId & """ data-optionid=""" & options(i).Sequence & """>" & vbCrLf

                ' text cell
                .Write "<td class=""dragHandle""><img alt="""" src=""../../../content/images/draghandle.png""/></td>" & vbCrLf
                .Write "<td>" & vbCrLf
                if selected then
                    .Write "<input type=""text"" name=""txtText"" id=""txtText"" class=""fullWidth"" maxlength=""30"" value=""" & options(i).Text & """/>" & vbCrLf
                else
                    .Write options(i).Text
                end if
                .Write "</td>" & vbCrLf

                ' value cell
                if showValueColumn then
                    .Write "<td>" & vbCrLf
                else
                    .Write "<td style=""display: none"">" & vbCrLf
                end if

                dim valueText
                if isnull(options(i).value) then
                    valueText = ""
                else
                    valueText = FormatNumber(options(i).value, 2)
                end if

                if selected then
                    .Write "<input type=""text"" name=""txtValue"" id=""txtValue"" class=""fullWidth"" maxlength=""100"" value=""" & valueText & """/>" & vbCrLf
                    .Write "<br/>" & vbCrLf
                    .Write "<span id=""valValueFormat"" style=""color: red; display: none"">Invalid number</span>" & vbCrLf
                else
                    .Write valueText
                end if
                .Write "</td>" & vbCrLf

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

            .Write "</tbody>" & vbCrLf

            ' show the footer row
            if len(editSequence) = 0 then
                .Write "<tr>" & vbCrLf
                .Write "<td/>" & vbCrLf
                .Write "<td>" & vbCrLf
                .Write "<input type=""text"" name=""txtText"" id=""txtText"" class=""fullWidth"" maxlength=""100"" value=""""/>" & vbCrLf
                .Write "</td>" & vbCrLf
                if showValueColumn then
                    .Write "<td>" & vbCrLf
                else
                    .Write "<td style=""display: none"">" & vbCrLf
                end if
                .Write "<input type=""text"" name=""txtValue"" id=""txtValue"" class=""fullWidth"" maxlength=""100"" value=""""/>" & vbCrLf
                .Write "<br/>" & vbCrLf
                .Write "<span id=""valValueFormat"" style=""color: red; display: none"">Invalid number</span>" & vbCrLf
                .Write "</td>" & vbCrLf
                .Write "<td>" & vbCrLf
                .Write "<a href=""javascript: add();"">Add</a>" & vbCrLf
                .Write "</td>" & vbCrLf
                .Write "</tr>" & vbCrLf
            end if
        end with
    end sub
%>