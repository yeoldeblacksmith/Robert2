<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title></title>
        <link type="text/css" rel="stylesheet" href="../content/css/jquery.ui.all.css" />

        <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.ui.core.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.ui.widget.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.ui.datepicker.min.js"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                var today = new Date;
                var formattedDate = (today.getMonth() + 1) + "/" + today.getDate() + "/" + today.getFullYear();

                var dateField = $("#txtPlayDate");

                dateField.val(formattedDate);
                dateField.datepicker({
                    onSelect: function (dateText, inst) { loadTimes(dateText); }
                });

                loadTimes(formattedDate);
            });

            function loadTimes(selectedDate) {
                $.ajax({
                    url: '../ajax/availability.asp?av=2&dt=' + selectedDate,
                    error: function () { alert("problems encountered"); },
                    success: function (data) {
                        $("#ddlPlayTime").html(data);
                    }
                });
            }

            function updateParent() {
                parent.$("#txtPlayDate").val($("#txtPlayDate").val());
                parent.$("#txtPlayTime").val($("#ddlPlayTime").val());
                parent.$.fancybox.close();
            }
        </script>
    </head>

    <body>
        <p style="text-align: center">
            <h2 align="center">What Day and Time Will You Be Playing?</h2>
        </p>

        <table style="margin: 0 auto" >
            <tr>
                <td>Date:</td>
                <td>
                    <input type="text" id="txtPlayDate" /><br />
                </td>
            </tr>
            <tr>
                <td>Time:</td>
                <td>
                    <select id="ddlPlayTime" />
                </td>
            </tr>
            <tr>
                <td colspan="2" style="text-align: center">
                    <button onClick="updateParent()">Ok</button>
                </td>
            </tr>
        </table>
    </body>
</html>
