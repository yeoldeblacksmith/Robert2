<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title></title>

        <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                loadMonthList();
                loadDayList();
                loadYearList();
            });

            function form_onSubmit() {
                if(validateDOB() & validateEmail()){
                    $.ajax({
                        url: '../ajax/waiver.asp?act=10&em=' + $("#txtEmail").val() + "&dt=" + $("#ddlDOBMonth").val() + "/" + $("#ddlDOBDay").val() + "/" + $("#ddlDOBYear").val(),
                        error: function () { alert("problems encountered"); },
                        success: function (data) {
                            if (data.length == 0) {
                                $("#valNotFound").css("display", "block");
                            } else {
                                parent.showLookupSuccess();
                            }
                        }
                    });
                }

                return false;
            }

            function loadDayList(selectedDay) {
                for (i = 1; i < 32; i++)
                    $("#ddlDOBDay").append(new Option(i, i, false, false));
            }

            function loadMonthList() {
                for (i = 1; i < 13; i++)
                    $("#ddlDOBMonth").append(new Option(i, i, false, false));
            }

            function loadYearList() {
                var today = new Date();
                var minYear = today.getFullYear() - 90;

                for (i = today.getFullYear() ; i >= minYear; i--)
                    $("#ddlDOBYear").append(new Option(i, i, false, false));
            }

            function validateDOB() {
                var valid = true;

                if ($("#ddlDOBMonth").val() == "" || $("#ddlDOBDay").val() == "" || $("#ddlDOBYear").val() == "") {
                    valid = false;
                    $("#valDOBReq").css("display", "block");
                } else {
                    $("#valFirstNameReq").css("display", "none");
                }

                return valid;
            }

            function validateEmail() {
                var valid = true;
                
                if ($("#txtEmail").val().length == 0) {
                    valid = false;
                    $("#valEmailReq").css("display", "block");
                } else {
                    $("#valEmailReq").css("display", "none");
                }

                return valid;
            }
        </script>
    </head>

    <body>
        <form>
            <h2 style="text-align: center">Waiver Lookup</h2>
            <p style="text-align: center">To help us locate your waiver, please enter the following information.</p>

            <table style="margin: 0 auto">
                <tr>
                    <td>
                        Email:
                    </td>
                    <td>
                        <input type="text" id="txtEmail" />
                        <span style="color: red; display: none" id="valEmailReq">Required</span>
                    </td>
                </tr>
                <tr>
                    <td>
                        Date of birth:
                    </td>
                    <td>
                        <select id="ddlDOBMonth" name="ddlDOBMonth" style="width: 50px"><option value="">MM</option></select>
                        <select id="ddlDOBDay" name="ddlDOBDay" style="width: 50px"><option value="">DD</option></select>
                        <select id="ddlDOBYear" name="ddlDOBYear" style="width: 65px"><option value="">YYYY</option></select><br />
                        <span style="color: red; display: none" id="valDOBReq">Required</span>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" style="text-align: center">
                        <button type="button" onclick="form_onSubmit()">Submit</button>
                    </td>
                </tr>
                <tr>
                    <td colspan="2" style="text-align: center">
                        <span style="color: red; display: none" id="valNotFound">Waiver not found. Please try again.</span>
                    </td>
                </tr>
            </table>
        </form>
    </body>
</html>
