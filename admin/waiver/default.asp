<!DOCTYPE html>
<!--#include file="../../classes/includelist.asp"-->

<%
    ' authenticate and authorize using pl/securityhelper
    AuthorizeForSiteUsers

    Dim useSelectedDate
    Dim highlightThisRow: highlightThisRow = ""
    Dim callSelectEventJS: callSelectEventJS = ""

    If IsNULL(request.querystring(QUERYSTRING_VAR_SELECTEDDATE)) OR IsEmpty(request.querystring(QUERYSTRING_VAR_SELECTEDDATE)) OR Len(request.querystring(QUERYSTRING_VAR_SELECTEDDATE)) = 0 Then
        useSelectedDate = "(today.getMonth() + 1) + ""/"" + today.getDate() + ""/"" + today.getFullYear();"
    Else
        useSelectedDate =  """" & Mid(request.querystring(QUERYSTRING_VAR_SELECTEDDATE),6,2) & "/" & Right(request.querystring(QUERYSTRING_VAR_SELECTEDDATE),2) & "/" & Left(request.querystring(QUERYSTRING_VAR_SELECTEDDATE),4) & """"
    End If

    If NOT (IsNULL(request.querystring(QUERYSTRING_VAR_ID)) OR IsEmpty(request.querystring(QUERYSTRING_VAR_ID)) OR Len(request.querystring(QUERYSTRING_VAR_ID)) = 0) Then
        callSelectEventJS = "selectEvent(" & request.querystring(QUERYSTRING_VAR_ID) & ");" & vbcrlf
        highlightThisRow = "highlightSelectedGroup(""myLine" & request.querystring(QUERYSTRING_VAR_ID) & """);" & vbcrlf
    End If

 
%>
<!--[if lt IE 7]>      <html class="ie6 lt-ie9 lt-ie8 lt-ie7" xmlns="http://www.w3.org/1999/xhtml"> <![endif]-->
<!--[if IE 7]>         <html class="ie7-js lt-ie9 lt-ie8" xmlns="http://www.w3.org/1999/xhtml"> <![endif]-->
<!--[if IE 8]>         <html class="ie8 lt-ie9" xmlns="http://www.w3.org/1999/xhtml"> <![endif]-->
<!--[if gt IE 8]><!--> <html class="default" xmlns="http://www.w3.org/1999/xhtml"> <!--<![endif]-->
    <head>
        <title><%= SiteInfo.Name %>  - Waivers by day</title>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">

        <link type="text/css" rel="Stylesheet" href="../../content/css/eventregistration.css?<%= ANTI_CACHE_STRING %>" />
        <link type="text/css" rel="stylesheet" href="../../content/css/jquery.ui.all.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/jquery.fancybox-1.3.4.css" />
        <link type="text/css" rel="Stylesheet" href="../../content/css/jquery.tablesorter.css" />
        <link type="text/css" rel="stylesheet" href="../../content/css/admin.css?<%= ANTI_CACHE_STRING %>" />
        <link type="text/css" rel="stylesheet" href="../../content/css/navmenu.css?<%= ANTI_CACHE_STRING %>" />
        <link type="text/css" rel="stylesheet" href="../../content/css/jquery.powertip-light.css" />
                
        <script type="text/javascript" src="../../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.ui.core.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.ui.widget.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.ui.datepicker.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.tablesorter.min.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.fancybox-1.3.4.pack.js"></script>
        <script type="text/javascript" src="../../content/js/jquery.powertip.min.js"></script>
        <script type="text/javascript">

            var currentlyHighlightedLine = "";

            $(document).ready(function () {
                var today = new Date;
                var formattedDate = <%= useSelectedDate %>

                var dateField = $("#txtSelectedDate");
                dateField.val(formattedDate);
                dateField.datepicker({
                    onSelect: function (dateText, inst) { selectDay(dateText); }
                });

                selectDay(formattedDate);

                dateField.keyup(function (e) {
                    if (e.keyCode == 13) {
                        selectDay(dateField.val());
                        dateField.datepicker('hide');
                    }
                });

                $('#txtSearch').keyup(function (e) {
                    if (e.keyCode == 13) {
                        search();
                    }
                });

                $(".waiverTable").css("height:", "100%");

            });

            function addPlayDate(playerId, waiverid, act) {
                var now = new Date();
                showPlayDate(act, playerId, (now.getMonth() + 1) + '/' + now.getDate() + '/' + now.getFullYear(), '00:00:00',null,null,null,null);
                selectWaiver(waiverid, $("#txtSelectedDate").val());
            }

            function checkInMulti(FromWhatPanel) {
                    var formvalues, noCheckedPlayerFound, msg;
                    noCheckedPlayerFound = true;
                    formvalues = jQuery("#myForm").serializeArray();

                    if (formvalues.length != 0) {
                        for (var i = 0; i < formvalues.length; i++) {
                            if (formvalues[i].name.substr(0, 12) == "txtCheckMeIn") {
                                if ($("#" + formvalues[i].name).attr('checked')) {
                                    noCheckedPlayerFound = false;
                                    $.ajax({
                                        url: '../../ajax/Waiver.asp?act=60&' + formvalues[i].value,
                                        error: function () { alert("problems encountered\nError Code: chkIMP"); },
                                        success: function (data) {
                                            var inputDate = new Date($("#txtSelectedDate").val());
                                            var todaysDate = new Date();
                                            if (inputDate.setHours(0, 0, 0, 0) != todaysDate.setHours(0, 0, 0, 0)) {
                                                msg = data;
                                            }
                                        },
                                        async: false
                                    });
                                }
                            }
                        }
                    }

                    if (noCheckedPlayerFound) {
                        alert("No players selected.");
                    } else {
                        if (msg && (FromWhatPanel != 'search')) {
                            alert("Player check in complete.");
                        }
                        refreshWaiverListPanel();
                    }
                }

            function checkInPlayer(FromWhatPanel) {
                var querystringvalues = "";
                var tmppid = $("#PlayerId").val();

                    querystringvalues = "&pd=" + $("#PlayerId").val() + "&pdt=" + $("#WaiverDate").val() + "&pt=" + $("#txtSelectedPlayTime").val();
                    //alert('../../ajax/Waiver.asp?act=60' + querystringvalues);
                    $.ajax({
                        url: '../../ajax/Waiver.asp?act=60' + querystringvalues,
                        error: function () { alert("problems encountered\nError Code: chkIP1"); },
                        success: function (data) {
                            var inputDate = new Date($("#txtSelectedDate").val());
                            var todaysDate = new Date();
                            if (inputDate.setHours(0, 0, 0, 0) != todaysDate.setHours(0, 0, 0, 0)) {
                                if (data) {
                                    if (FromWhatPanel != 'search') {
                                        alert("Player check in complete");
                                    }
                                }
                            }
                        },
                        async: false
                    });

                    refreshWaiverListPanel();

                    if ($("#WaiverId").val() != "" && (typeof $("#WaiverId").val() != "undefined")) {
                        selectWaiver($("#WaiverId").val(), $("#WaiverDate").val());
                        $("#PlayerId").val(tmppid);
                        $('#WaiverFrame').css('display', 'block');
                    }
                }

            function checkOutPlayer(FromWhatPanel) {
               var querystringvalues = "";
               var tmppid = $("#PlayerId").val();
                    querystringvalues = "&pd=" + $("#PlayerId").val() + "&pdt=" + $("#WaiverDate").val() + "&pt=" + $("#txtSelectedPlayTime").val();
                    $.ajax({
                        url: '../../ajax/Waiver.asp?act=61' + querystringvalues,
                        error: function () { alert("problems encountered\nError Code: chkOP1"); },
                        success: function (data) {
                            var inputDate = new Date($("#txtSelectedDate").val());
                            var todaysDate = new Date();
                            if (inputDate.setHours(0, 0, 0, 0) != todaysDate.setHours(0, 0, 0, 0)) {
                                if (data) {
                                    if (FromWhatPanel != 'search') {
                                        alert("Player check out complete");
                                    }
                                }
                            }
                        },
                        async: false
                    });

                    refreshWaiverListPanel();

                    if ($("#WaiverId").val() != "" && (typeof $("#WaiverId").val() != "undefined")) {
                        selectWaiver($("#WaiverId").val(), $("#WaiverDate").val());
                        $("#PlayerId").val(tmppid);
                        $('#WaiverFrame').css('display', 'block');
                    }
                }

            function checkOutMulti(FromWhatPanel) {
                var formvalues, noCheckedPlayerFound, msg;
                msg = "";
                noCheckedPlayerFound = true;
                formvalues = jQuery("#myForm").serializeArray();

                if (formvalues.length != 0) {
                    for (var i = 0; i < formvalues.length; i++) {
                        if (formvalues[i].name.substr(0, 12) == "txtCheckMeIn") {
                            if ($("#" + formvalues[i].name).attr('checked')) {
                                noCheckedPlayerFound = false;
                                $.ajax({
                                    url: '../../ajax/Waiver.asp?act=61&' + formvalues[i].value,
                                    error: function () { alert("problems encountered\nError Code: chkOM1"); },
                                    success: function (data) {
                                        var inputDate = new Date($("#txtSelectedDate").val());
                                        var todaysDate = new Date();
                                        if (inputDate.setHours(0, 0, 0, 0) != todaysDate.setHours(0, 0, 0, 0)) { msg += data; }
                                    },
                                    async: false
                                });
                            }
                        }
                    }
                }

                if (noCheckedPlayerFound) { 
                    alert("No players selected.");
                } else {
                    if (msg != "" && FromWhatPanel != 'search') {
                        alert(msg);
                    }
                    refreshWaiverListPanel();
                }
            }

            function deletePlayDate(playerId, playdate, playtime, waiverid) {
                if (confirm("Are you sure you want to delete this play date?")) {
                    $.ajax({
                        url: '../../ajax/Waiver.asp?act=62&pd=' + playerId + '&pdt=' + playdate + '&pt=' + playtime,
                        error: function () { alert("problems encountered\nError Code: delPdt1"); },
                        success: function () { },
                        async: false
                    });

                    var todaysDate = new Date();
                    var selectedDate = new Date($("#txtSelectedDate").val());
                    var originalPlayDate = new Date(playdate);

                    if (todaysDate.setHours(0, 0, 0, 0) == originalPlayDate.setHours(0, 0, 0, 0)) {
                        refreshAllPanels();
                    } else {
                        if (selectedDate.setHours(0, 0, 0, 0) == originalPlayDate.setHours(0, 0, 0, 0)) {
                            $("#WaiverId").val("");
                            $("#WaiverDate").val("");
                            refreshAllPanels();
                        } else {
                            refreshWaiverPanel();
                        }
                    }
                }
            }

            function deleteWaiver() {
                if (confirm("Do you really want to delete this waiver?")) {
                    $.ajax({
                        url: '../../ajax/Waiver.asp?act=30&id=' + $("#WaiverId").val(),
                        error: function () { alert("problems encountered\nError Code: delW1 "); },
                        success: function () {
                            $("#WaiverId").val("");
                            $("#WaiverDate").val("");
                            refreshAllPanels();
                        },
                        async: false
                    });
                }
            }

            function highlightSelectedGroup(whichLineItem){
                if (whichLineItem == undefined){
                    whichLineItem = currentlyHighlightedLine;
                }

                $("#" + currentlyHighlightedLine).css('background','');
                currentlyHighlightedLine = whichLineItem;
                $("#" + currentlyHighlightedLine).css('background','#94FF70');
            }

            function refreshAllPanels() {
                refreshGroupListPanel();
                refreshWaiverListPanel();
                refreshWaiverPanel();
            }
            function refreshGroupListPanel() {
                $("#GroupList").empty();

                if ($("#txtSelectedDate").val() == "") {
                    $("#GroupList").html("Search Results");
                } else {
                    $.ajax({
                        url: '../../ajax/Waiver.asp?act=53&dt=' + $("#txtSelectedDate").val(),
                        error: function () { alert("problems encountered selecting event groups"); },
                        success: function (data) {
                            $("#GroupList").html(data);
                        },
                        async: false
                    });
                }
                highlightSelectedGroup(currentlyHighlightedLine);
            }

            function refreshWaiverListPanel() {
                var ev = $("#EventId").val();
                var pt = $("#txtSelectedPlayTime").val();
                var sc = $("#txtSearch").val();

                $("#WaiverList").empty();

                if (sc != "") {
                    search();
                } else {
                    if (ev != "") {
                        selectEvent($("#EventId").val());
                    } else {
                        if (pt != "" && (typeof pt != 'undefined')) {
                            selectWalkupGroup(pt);
                        } 
                    }
                }
            }

            function refreshWaiverPanel() {
                var wid = $("#WaiverId").val();
                var wdt = $("#WaiverDate").val();

                if (wid != "" && wdt != "" && (typeof wid != 'undefined')) {
                    selectWaiver(wid, wdt);
                } else {
                    $('#WaiverFrame').css('display', 'none');
                }
            }

            function resetPanels() {
                selectDay($("#txtSelectedDate").val());

                if ($("#GroupList").html().length == 0) {
                    $("#EventId").val("");
                    $("#WaiverList").empty();
                } else {
                    if ($("#EventId").val().length == 0) {
                        $("#WaiverList").empty();
                    }else {
                        selectEvent($("#EventId").val());
                    }
                }

                $('#WaiverFrame').css('display', 'none');
            }

            function selectDay(SelectedDate) {
                $("#txtSearch").val("");

                $.ajax({
                    url: '../../ajax/Waiver.asp?act=53&dt=' + SelectedDate,
                    error: function () { alert("problems encountered selecting event groups"); },
                    success: function (data) {
                        $("#GroupList").html(data);
                        $("#EventId").val("");
                        $("#PlayerId").val("");
                        $("#txtSelectedPlayTime").val("");
                        $("#WaiverList").empty();
                        $('#WaiverFrame').css('display', 'none');
                        $("#txtPlayDate").val(SelectedDate);
                    },
                    async: false
                });

                <%= callSelectEventJS %>
                <%= highlightThisRow %>          

            }

            function selectEvent(EventId,whichLineItem) {
                highlightSelectedGroup(whichLineItem);

                //alert('../../ajax/Waiver.asp?act=54&ev=' + EventId + '&dt=' + $("#txtSelectedDate").val());
                $.ajax({
                    url: '../../ajax/Waiver.asp?act=54&ev=' + EventId + '&dt=' + $("#txtSelectedDate").val(),
                    error: function () { alert("problems encountered\nError Code: selEv1"); },
                    success: function (data) {
                        $("#WaiverList").html(data);
                        $("#EventId").val(EventId);
                        //$("#txtSelectedPlayTime").val("");
                        $("#PlayerId").val("");
                        $('#WaiverFrame').css('display', 'none');
                        $(".results").tablesorter(
                            {
                            // pass the headers argument and assing a object 
                            headers: {
                            //    // assign the secound column (we start counting zero) 
                                0: {
                                    // disable it by setting the property sorter to false 
                                    sorter: false
                                   }
                            }
                            }
                        );
                    },
                    async: false
                });
            }

            function selectWalkupGroup(PlayTime,whichLineItem) {
                highlightSelectedGroup(whichLineItem);
                
                //alert('../../ajax/Waiver.asp?act=55&pt=' + PlayTime + '&dt=' + $("#txtSelectedDate").val());
                $.ajax({
                    url: '../../ajax/Waiver.asp?act=55&pt=' + PlayTime + '&dt=' + $("#txtSelectedDate").val(),
                    error: function () { alert("problems encountered\nError Code: selWU1"); },
                    success: function (data) {
                        $("#WaiverList").html(data);
                        $("#EventId").val("");
                        $("#txtSelectedPlayTime").val(PlayTime);
                        $("#PlayerId").val("");
                        $('#WaiverFrame').css('display', 'none');

                        $(".results").tablesorter({ 
                            // pass the headers argument and assing a object 
                            headers: { 
                                // assign the secound column (we start counting zero) 
                                0: { 
                                    // disable it by setting the property sorter to false 
                                    sorter: false 
                                } 
                            } 
                        });
                    },
                    async: false
                });
            }

            function selectWaiver(WaiverId, WaiverDate) {
                $("#PlayerId").val("");

                var frame = $('#WaiverFrame');
                frame.prop('src', 'display.asp?id=' + WaiverId + "&dt=" + $("#txtSelectedDate").val());
                frame.css('display', 'block');

                $("#WaiverId").val(WaiverId);
                $("#WaiverDate").val(WaiverDate);

            }

            function selectAllPlayers(flag) {
                var values = jQuery("#myForm").serializeArray();
                    values = values.concat(
                        jQuery('#myForm input[type=checkbox]:not(:checked)').map(
                            function () {
                                return { "name": this.name, "value": this.value }
                        }).get()
                    );

                    if (values.length != 0) {
                        for (var i = 0; i < values.length; i++) {
                            if (values[i].name.substr(0, 12) == "txtCheckMeIn") {
                            if (flag == 'Select All') {
                                $("#" + values[i].name).attr('checked', true);
                            } else {
                                $("#" + values[i].name).attr('checked', false);
                            }
                        }
                    }
                }
            }

            function search() {
                $('#WaiverFrame').css('display', 'none');

                $.ajax({
                    url: '../../ajax/Waiver.asp?act=56&tx=' + escape($("#txtSearch").val()),
                    error: function () { alert("problems encountered\nError Code: srch1"); },
                    success: function (data) {
                        $("#GroupList").html("Search Results");
                        $("#WaiverList").html(data);
                    }
                });

                $("#txtSelectedDate").val("");
                $("#EventId").val("");
                $("#txtSelectedPlayTime").val("");
            }

            function showEventGroups(playerid, playdate, playtime) {
                $.fancybox({
                    'type': 'iframe',
                    'href': 'eventgroup.asp?pd=' + playerid + '&pdt=' + playdate + '&pt=' + playtime,
                    'titlePosition': 'outside',
                    'title': '&copy;Vantora',
                    'width': 350,
                    'height': 275,
                    'transitionIn': 'elastic',
                    'transitionOut': 'elastic',
                    'speedIn': 200,
                    'speedOut': 200
                });
            }

            function showEventGroupsForMultiselect() {
                var formvalues, querystringvalues, noCheckedPlayerFound;
                noCheckedPlayerFound = true;
                querystringvalues = "";
                formvalues = jQuery("#myForm").serializeArray();

                if (formvalues.length != 0) {
                    for (var i = 0; i < formvalues.length; i++) {
                        if (formvalues[i].name.substr(0, 12) == "txtCheckMeIn") {
                            noCheckedPlayerFound = false;
                            if ($("#" + formvalues[i].name).attr('checked')) {
                                querystringvalues += "@" + formvalues[i].value.substr(3, (formvalues[i].value.indexOf('&') - 3));
                            }
                        }
                    }

                    if (noCheckedPlayerFound) {
                        alert("No players selected.");
                    } else {
                        $.fancybox({
                            'type': 'iframe',
                            'href': 'eventgroup.asp?pi=' + querystringvalues + '&pdt=' + $("#txtSelectedDate").val() + '&pt=' + $("#txtSelectedPlayTime").val(),
                            'titlePosition': 'outside',
                            'title': '&copy;Vantora',
                            'width': 350,
                            'height': 275,
                            'transitionIn': 'elastic',
                            'transitionOut': 'elastic',
                            'speedIn': 200,
                            'speedOut': 200
                        });
                    }
                } else { alert("No players selected."); }
            }

            function setCheckedInFilter() {
                if ($("#cbxIncCheckedIn").attr('checked')) {
                    document.cookie = "IncCheckedIn=true; expires=Thu, 2 Aug 2021 20:47:11 UTC; path=/";
                } else {
                    document.cookie = "IncCheckedIn=false; expires=Thu, 2 Aug 2021 20:47:11 UTC; path=/";
                }

                refreshWaiverListPanel();
            }

            function setExpiredFilter() {
                if ($("#cbxIncExpired").attr('checked')) {
                    document.cookie = "IncExpired=true; expires=Thu, 2 Aug 2021 20:47:11 UTC; path=/";
                } else {
                    document.cookie = "IncExpired=false; expires=Thu, 2 Aug 2021 20:47:11 UTC; path=/";
                }
                refreshWaiverListPanel();
            }

            function showPlayDate(act, playerid, playdate, playtime, eventid, checkintime, originalplaydate, originalplaytime) {
                $.fancybox({
                    'type' : 'iframe',
                    'href': 'playdate.asp?act=' + act + '&pd=' + playerid + '&pdt=' + playdate + '&pt=' + playtime + "&ev=" + eventid + "&aci=" + checkintime + "&opdt=" + originalplaydate + "&opt=" + originalplaytime,
                    'titlePosition': 'outside',
                    'title': '&copy;Vantora',
                    'width': 350,
                    'height': 275,
                    'transitionIn': 'elastic',
                    'transitionOut': 'elastic',
                    'speedIn': 200,
                    'speedOut': 200
                });
            }
         </script>
    </head>

    <body>
<%
    dim navmenu
    set navmenu = new NavigationMenu
    navmenu.HeadingTitle = "Waiver Administration"

    navmenu.FormControlStrings.Add "<input type=""text"" id=""txtSearch"" maxlength=""1000"" tabindex=""3"" />"
    navmenu.FormControlStrings.Add "<img src=""../../content/images/magnifier.png"" alt="""" title=""Search"" onclick=""search()"" tabindex=""4"" />"
    navmenu.FormControlStrings.Add "<input type=""text"" id=""txtSelectedDate"" style=""width: 75px"" tabindex=""1"" />"
    navmenu.FormControlStrings.Add "<img src=""../../content/images/calendar_view_month.png"" alt="""" title=""date"" onclick=""$('#txtSelectedDate').datepicker('show');"" tabindex=""2"" />"

    navmenu.WriteNavigationSection NAVIGATION_NAME_WAIVERS
%>
        <input id="EventId" type="hidden" value="" />
        <input id="WaiverId" type="hidden" value="" />
        <input id="PlayerId" type="hidden" value="" />
        
        <form id="myForm" style="margin: 0; height:100%">   
            <table class="adminTable waiverTable" height="100%">
                <tr>
                    <td class="adminCell" id="GroupList" style="vertical-align: top; width: 200px"></td>
                    <td class="adminCell" id="WaiverList" style="vertical-align: top; width: 400px"></td>
                    <td class="adminCell" id="Waiver"  style="height: 100%; vertical-align: top;">
                        <input type="hidden" id="WaiverDate" />

                        <iframe id="WaiverFrame" frameborder="0" draggable="false" style="display: none; height: 95%; width: 99%; border: 0px"></iframe>
                    </td>
                </tr>
            </table>
        </form>
    </body>
</html>
