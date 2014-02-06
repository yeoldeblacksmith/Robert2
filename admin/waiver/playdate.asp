<!DOCTYPE html>
<%
    CONST ACTION_UPDATEPLAYDATE = "1"
    CONST ACTION_ADDNEWPLAYDATE = "2"
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title></title>
        <link type="text/css" rel="stylesheet" href="../../content/css/jquery.ui.all.css" />

        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.ui.core.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.ui.widget.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.ui.datepicker.min.js"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                var playDate = $("#txtPlayDate");
                playDate.val('<%= Request.Querystring("pdt") %>');
                playDate.datepicker();
            });

            function clearMessage(){
                $("#txtMsg").html("");
            }

            function updatePlayDate() {
                $.ajax({
<%
    Select case request.QueryString("act")
        Case ACTION_UPDATEPLAYDATE
%>
                    url: '../../ajax/Waiver.asp?act=57&pd=<%= Request.Querystring("pd") %>&pdt=' + $("#txtPlayDate").val() + '&pt=<%= Request.Querystring("pt") %>&ev=<%= Request.Querystring("ev") %>&aci=&opdt=<%= Request.Querystring("pdt") %>&opt=<%= Request.Querystring("opt") %>',
<%
        Case ACTION_ADDNEWPLAYDATE
%>
                    url: '../../ajax/Waiver.asp?act=59&pd=<%= Request.Querystring("pd") %>&pdt=' + $("#txtPlayDate").val() + '&pt=<%= Request.Querystring("pt") %>',
<%
    End Select
%>
                    error: function () { alert("problems encountered"); },
                    success: function (data) {
                        if (data.length != 0){
                            $("#txtMsg").html(data);
                        }else{
                            var selectedDate = new Date(parent.$("#txtSelectedDate").val());
                            var originalPlayDate = new Date($("#txtOriginalPlayDate").val());

                            if (selectedDate.setHours(0, 0, 0, 0) == originalPlayDate.setHours(0, 0, 0, 0)) {
                                parent.$("#WaiverId").val("");
                                parent.$("#WaiverDate").val("");
                                parent.refreshAllPanels();
                            }else {
                                parent.refreshAllPanels();
                            }
                            
                            parent.$.fancybox.close();
                        }
                    },
                    async: false
                });
             }
        </script>
    </head>

    <body>
            <div id="divPlayDate" class="center">
                <p>Please choose another date<Br />
                     for the selected player:</p>
                <input type="text" id="txtPlayDate" onChange="clearMessage();" />
                <br />
                <span id="txtMsg" style="color: red;"></span><br /><br />
                <button onclick="updatePlayDate()">Update</button>
                <button onclick="parent.$.fancybox.close();">Cancel</button>
                <input type="hidden" id="txtOriginalPlayDate" name="txtOriginalPlayDate" value="<%= Request.Querystring("pdt") %>" />
            </div>
    </body>
</html>
