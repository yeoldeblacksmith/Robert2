<!DOCTYPE html>
<!--#include file="../../classes/IncludeList.asp" -->
<%
    AuthorizeForAdminOnly

    if Request.ServerVariables("REQUEST_METHOD") = "POST" then

        dim waivers : set waivers = new WaiverCollection
        dim myCustomFieldsColl: set myCustomFieldsColl = new WaiverCustomFieldCollection
            myCustomFieldsColl.LoadDistinct EXCLUDEINACTIVEWAIVERCUSTOMFIELDS
        dim arrColumnNames: arrColumnNames = Array()
        dim arrColumnValues: arrColumnValues = Array()
        ReDim arrColumnValues(1,0)

        dim previousFlag : previousFlag = cbool(request.Form("chkPrevious") = "on")
        dim fileName : fileName =   replace(SiteInfo.Name, " ", "_") & _
                                    "-" & request.Form("radOption") & "-" & _
                                    DatePart("yyyy", Now()) & _
                                    iif(DatePart("m", Now()) < 10, "0" & DatePart("m", Now()), DatePart("m", Now())) & _
                                    iif(DatePart("d", Now()) < 10, "0" & DatePart("d", Now()), DatePart("d", Now())) & _ 
                                    iif(DatePart("h", Now()) < 10, "0" & DatePart("h", Now()), DatePart("h", Now())) & _
                                    iif(DatePart("n", Now()) < 10, "0" & DatePart("n", Now()), DatePart("n", Now())) & _
                                    iif(DatePart("s", Now()) < 10, "0" & DatePart("s", Now()), DatePart("s", Now())) & _
                                    ".csv"

        select case request.Form("radOption")
            case "all"
                waivers.LoadForExportByAll previousFlag
            case "custom"
                waivers.LoadForExportByCustomRange request.Form("txtStartDate"), request.Form("txtEndDate"), previousFlag
            case "month"
                waivers.LoadForExportByMonth previousFlag
            case "year"
                waivers.LoadForExportByYear previousFlag
        end select

        response.Clear()
        Response.ContentType = MIMETYPE_OTHER
        Response.AddHeader "content-disposition", "attachment; filename=" & fileName
        response.Buffer = true
    
        if SiteInfo.Country = "US" then
            response.write """WaiverId"",""FirstName"",""LastName"",""Age"",""DateOfBirth"",""ParentFirstName"",""ParentLastName"",""PhoneNumber"",""Address"",""City"",""State"",""ZipCode""," & addCustomFieldColumnNames(WAIVERCUSTOMFIELDS_LOCATION_ADULTANDMINORS) & addCustomFieldColumnNames(WAIVERCUSTOMFIELDS_LOCATION_ADULTONLY) & addCustomFieldColumnNames(WAIVERCUSTOMFIELDS_LOCATION_MINORSONLY) & addCustomFieldColumnNames(WAIVERCUSTOMFIELDS_LOCATION_WAIVERGENERAL) & """PlayDate"",""PlayTime"",""CreateDateTime"",""EmailAddress"",""EmailList""" & vbCrLf
        else
            response.Write """WaiverId"",""FirstName"",""LastName"",""Age"",""DateOfBirth"",""ParentFirstName"",""ParentLastName"",""PhoneNumber"",""Address"",""City"",""Province"",""PostalCode""," & addCustomFieldColumnNames(WAIVERCUSTOMFIELDS_LOCATION_ADULTANDMINORS) & addCustomFieldColumnNames(WAIVERCUSTOMFIELDS_LOCATION_ADULTONLY) & addCustomFieldColumnNames(WAIVERCUSTOMFIELDS_LOCATION_MINORSONLY) & addCustomFieldColumnNames(WAIVERCUSTOMFIELDS_LOCATION_WAIVERGENERAL) & """PlayDate"",""PlayTime"",""CreateDateTime"",""EmailAddress"",""EmailList""" & vbCrLf
        end if

        for i = 0 to waivers.Count - 1
            response.Write waivers(i).ToString() & vbCrLf
            response.Flush
        next    

        waivers.MarkAsExported
        response.End()
    end if
%>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title>Waiver Export</title>
        <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/admin.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/navmenu.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/jquery.ui.all.css" />

        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.ui.core.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.ui.widget.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.ui.datepicker.min.js"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                $("#txtEndDate").datepicker({ 
                    onSelect: function (dateText, instr) { $("#txtEndDate").val(dateText); }
                });

                $("#txtStartDate").datepicker({
                    onSelect: function (dateText, instr) { $("#txtStartDate").val(dateText); }
                });
            });

            function radOption_onclick(value) {
                if (value == "custom") {
                    $("#tblRange").css("display", "block");
                } else {
                    $("#tblRange").css("display", "none");
                }
            }

            function setEndDate(value) {
                $("#txtEndDate").val(value);
            }

            function setStartDate(value) {
                $("#txtStartDate").val(value);
            }

            function validate() {
                var valid = true;

                if ($("#radOptionCustom").is(":checked")) {
                    var startDate;
                    var endDate;

                    if ($("#txtStartDate").val().length == 0) {
                        $("#valStartDateReq").css("display", "inline");
                        valid = false;
                    } else {
                        $("#valStartDateReq").css("display", "none");
                        startDate = new Date($("#txtStartDate").val());
                    }

                    if ($("#txtEndDate").val().length == 0) {
                        $("#valEndDateReq").css("display", "inline");
                        valid = false;
                    } else {
                        $("#valEndDateReq").css("display", "none");
                        endDate = new Date($("#txtEndDate").val());
                    }

                    // if the dates are swapped
                    if (typeof startDate != 'undefined' && typeof endDate != 'undefined') {
                        if (startDate > endDate) {
                            $("#txtStartDate").val(endDate.toLocaleDateString());
                            $("#txtEndDate").val(startDate.toLocaleDateString());
                        }
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
    navmenu.HeadingTitle = "Waiver Export"

    navmenu.WriteNavigationSection NAVIGATION_NAME_WAIVERS 
%>
        <form method="post">
            <table class="adminTable" style="margin: 0 auto; width: 550px; height: auto;">
                <tr>
                    <td class="adminCell" style=" padding: 10px">
                        <p style="font-weight: bold">Export Options:</p>

                        <div>
                            <input type="radio" name="radOption" id="radOptionAll" value="all" onchange="return radOption_onclick(this.value)" <%= iif(len(request.Form("radOption")) = 0 or request.Form("radOption") = "all", "checked=""checked""", "") %> />
                            <label for="radOptionAll">All</label><br />

                            <input type="radio" name="radOption" id="radOptionYear" value="year" onchange="return radOption_onclick(this.value)" <%= iif(request.Form("radOption") = "year", "checked=""checked""", "") %> />
                            <label for="radOptionYear">Year</label><br />

                            <input type="radio" name="radOption" id="radOptionMonth" value="month" onchange="return radOption_onclick(this.value)" <%= iif(request.Form("radOption") = "month", "checked=""checked""", "") %> />
                            <label for="radOptionMonth">Month</label><br />
                            
                            <input type="radio" name="radOption" id="radOptionCustom" value="custom" onchange="return radOption_onclick(this.value)" <%= iif(request.Form("radOption") = "custom", "checked=""checked""", "") %> />
                            <label for="radOptionCustom">Custom</label><br />
                        </div>

                        <table id="tblRange" class="tableDefault" border="0"  <%= iif(request.Form("radOption") = "custom", "", "style=""display: none""") %>>
                            <tr>
                                <td>
                                    <label for="txtStartDate">From:</label>
                                </td>
                                <td>
                                    <input type="text" name="txtStartDate" id="txtStartDate" />
                                </td>
                                <td>
                                    <img src="../../content/images/calendar_view_month.png" alt="" title="Start Date" onclick="$('#txtStartDate').datepicker('show');" />
                                </td>
                            </tr>
                            <tr>
                                <td colspan="3">
                                    <span id="valStartDateReq" style="display: none; color: red">Required</span>
                                </td>
                            </tr>
                            <tr>
                                <td>
                                    <label for="txtEndDate">To:</label>
                                </td>
                                <td>
                                    <input type="text" name="txtEndDate" id="txtEndDate" />
                                </td>
                                <td>
                                    <img src="../../content/images/calendar_view_month.png" alt="" title="End Date" onclick="$('#txtEndDate').datepicker('show');" /><br />
                                </td>
                            </tr>
                            <tr>
                                <td colspan="3">
                                    <span id="valEndDateReq" style="display: none; color: red">Required</span>
                                </td>
                            </tr>
                        </table>

                        <button class="center" style="display: block" onclick="return validate()">Export</button>

                        <input type="checkbox" name="chkPrevious" id="chkPrevious" />
                        <label for="chkPrevious">Include previously exported records</label>
                    </td>
                </tr>
            </table>

        </form>
    </body>
</html>
<%
    function fillCustomFieldsColumns(column, stringToUnpack)
        dim retVal: retVal = ""
        dim str: str =  stringToUnpack
        dim currentPosition: currentPosition = 1
        dim eovalue: eovalue = 0
        dim eoname: eoname = 0
        redim arrColumnValues(1,0)

        If NOT (isNULL(str)) Then
            str = Right(str,Len(str)-eovalue)
            if len(str) <> 0 then
               do until Len(str) <= 0
                    currentPosition = 1
                    eoname = Instr(currentPosition,str,"^")-1
                    ReDim Preserve arrColumnValues(1,ubound(arrColumnValues,2)+1)
                    arrColumnValues(0,UBound(arrColumnValues,2)) = Mid(str,currentPosition,eoname)  
                    eovalue = Instr(currentPosition,str,"~")-1
                    if eovalue <= 0 then
                        eovalue = Len(str)
                    end if
                    currentPosition = eoname + 2
                    arrColumnValues(1,UBound(arrColumnValues,2)) = Mid(str,currentPosition,(eovalue-(currentPosition-1)))             
                    if NOT (Len(str)-(eovalue+1) <= 0) then
                        str = Right(str,Len(str)-(eovalue+1))
                    else
                        str = "" 
                    end if
                loop
            else
                retVal = ""  
            end if

            for m = 0 to Ubound(arrColumnValues,2)
                if arrColumnValues(0,m) = column then
                    retVal = arrColumnValues(1,m)
                    Exit For
                end if

            next
        else
        End if    

        fillCustomFieldsColumns = retVal
    end function

    function addCustomFieldColumnNames(location)
        dim retVal: retVal = ""
        for k = 0 to myCustomFieldsColl.Count-1
            if myCustomFieldsColl(k).FormLocation = location then
                retVal = retVal & """" & myCustomFieldsColl(k).Name & """," 
                ReDim Preserve arrColumnNames(Ubound(arrColumnNames)+1)
                arrColumnNames(UBound(arrColumnNames)) = myCustomFieldsColl(k).Name
             end if
        next
        addCustomFieldColumnNames = retVal
    end function
 %>