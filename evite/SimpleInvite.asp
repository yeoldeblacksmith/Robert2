<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
    <title></title>

    <script type="text/javascript" src="content/js/jquery-1.7.2.min.js"></script>
    <script type="text/javascript" src="content/js/invite.js"></script>
    <script type="text/javascript">
        $(document).ready(function () {
            loadTable();
        });

        function loadTable() {
            $.ajax({
                url: 'common/inviteajax.asp?act=20&id=' + $("#eventId").val(),
                failure: function () { alert("problems encountered loading table"); },
                success: function (data) {
                    var container = $("#datagrid");
                    
                    container.empty();
                    container.html(data);
                }
            });
        }

        function sendInvites() {
            $.ajax({
                url: 'common/inviteajax.asp?act=30&id=' + $("#eventId").val(),
                error: function () { alert("problems encountered"); },
                success: function () { parent.showSent(); }
            });
        }
    </script>
</head>
<body>
    <p style="text-align: center">Please enter your guest information</p>

    <input type="hidden" id="eventId" value="<%= Request.QueryString("id") %>" />

    <div id="datagrid"></div>

    <p style="text-align: center">
        <button type="button" id="btnSend" onclick="sendInvites()">Send Invites</button>
    </p>
</body>
</html>
