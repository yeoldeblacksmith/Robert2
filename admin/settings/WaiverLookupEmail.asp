<!DOCTYPE html>
<!--#include file="../../classes/IncludeList.asp" -->
<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForAdminOnly

    dim bSaved
    bSaved = false

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then
        for each key in Request.Form
            settings(key) = Request.Form(key)
            
            dim mySetting
            set mySetting = new SiteSetting
            mySetting.Key = key
            mySetting.Value = request.Form(key)
            mySetting.Save
        next

        bSaved = true
    end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Waiver Lookup Email Settings</title>
        <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/admin.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/navmenu.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/jquery.fancybox-1.3.4.css" />

        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.fancybox-1.3.4.pack.js"></script>
        <script type="text/javascript" src="../../content/js/admin.js"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                $("#templateLink").fancybox({
                    'type': 'iframe',
                    'href': 'WaiverEmailTemplate.html',
                    'titlePosition': 'outside',
                    'title': '&copy;Vantora',
                    'width': 500,
                    'height': 320,
                    'transitionIn': 'elastic',
                    'transitionOut': 'elastic',
                    'speedIn': 200,
                    'speedOut': 200
                });

                $("#previewEmail").click(function() {
                    previewType = "<%= SETTING_EMAILTEMPLATE_WAIVERLOOKUP %>"
                    previewEmail(previewType, $("#" + "<%= SETTING_EMAILSUBJECT_WAIVERLOOKUP %>").val(), $("#" + "<%= SETTING_EMAILTEMPLATE_WAIVERLOOKUP %>").val()); 
                });
            });

            function showUpload(parmName) {
                $.fancybox({
                    'type': 'iframe',
                    'href': 'templateupload.asp?nm=' + parmName,
                    'titlePosition': 'outside',
                    'title': '&copy;Vantora',
                    'width': 400,
                    'height': 75,
                    'transitionIn': 'elastic',
                    'transitionOut': 'elastic',
                    'speedIn': 200,
                    'speedOut': 200
                });
            }

            function validateForm() {
                var valid = true;

                return valid;
            }

        </script>
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Waiver Lookup Email Settings"

    navmenu.WriteNavigationSection NAVIGATION_NAME_SETTINGS 
%>
        <form method="post">
            <table class="adminTable" style="margin: 0 auto; width: 877px; height: auto;">
                <tr>
                    <td class="adminCell" style="padding: 10px">
                        <div class="floatRight">
                            <a id="templateLink" href="#" target="_blank">Field Templates</a>
                        </div>
                        <div class="clear floatNone"/>
                        <br />

                        <table class="tableBorder marginCenter fullWidth">
                            <tr class="alternate">
                                <td colspan="2">Subject</td>
                            </tr>
                            <tr>
                                <td colspan="2">
                                    <input name="<%= SETTING_EMAILSUBJECT_WAIVERLOOKUP%>" id="<%= SETTING_EMAILSUBJECT_WAIVERLOOKUP%>" value="<%= Settings(SETTING_EMAILSUBJECT_WAIVERLOOKUP) %>" class="fullWidth" />
                                </td>
                            </tr>
                            <tr class="alternate">
                                <td colspan="2">Body</td>
                            </tr>
                            <tr>
                                <td>Text:</td>
                                <td>
                                    <% 
                                        Dim bodyEditor
	                                    Set bodyEditor = New CuteEditor
                                        bodyEditor.FilesPath = "../../../../cute/CuteEditor_Files"
	                                    bodyEditor.HelpUrl = "help.asp"
                                        bodyEditor.ImageGalleryPath = "../../content/images/cuteupload"
                                        bodyEditor.ConfigurationPath = "../../../../cute/CuteEditor_Files/Configuration/autoconfigure/simple.config"
                                        bodyEditor.Id = SETTING_EMAILTEMPLATE_WAIVERLOOKUP
                                        bodyEditor.Text = settings(SETTING_EMAILTEMPLATE_WAIVERLOOKUP)
                                        bodyEditor.Draw()
                                    %>
                                </td>
                            </tr>
                        </table>

                        <br /><br /><br />

                        <%
                            if bSaved then
                        %>
                            <p class="center red">Settings saved successfully</p>
                        <%
                            end if
                        %>

                        <p id="valMessage" class="center red" style="display: none"></p>

                        <div class="marginCenter center">
                            <button type="submit" class="marginCenter" onClick="return validateForm()">Save Changes</button>
                            <button type="button" id="previewEmail">Preview</button>
                        </div>
                    </td>
                </tr>
            </table>
        </form>

    </body>
</html>
