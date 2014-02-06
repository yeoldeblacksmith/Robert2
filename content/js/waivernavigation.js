var currentMinor = -1;
var numberOfMinors = 0;
var selectedPath = -1;
var selectedSearch = -1;
var wizardStep = -1;
var checkInId = 0;

function buildDayList(DayControlId, SelectedMonth, SelectedYear) {
    var daysInMonth = 0;
    var currentDay = $("#" + DayControlId).val();

    switch (parseInt(SelectedMonth)) {
        case 1:
            daysInMonth = 31;
            break;
        case 2:
            if (SelectedYear % 4 == 0) {
                daysInMonth = 29;
            } else {
                daysInMonth = 28;
            }
            break;
        case 3:
            daysInMonth = 31;
            break;
        case 4:
            daysInMonth = 30;
            break;
        case 5:
            daysInMonth = 31;
            break;
        case 6:
            daysInMonth = 30;
            break;
        case 7:
            daysInMonth = 31;
            break;
        case 8:
            daysInMonth = 31;
            break;
        case 9:
            daysInMonth = 30;
            break;
        case 10:
            daysInMonth = 31;
            break;
        case 11:
            daysInMonth = 30;
            break;
        case 12:
            daysInMonth = 31;
            break;
    }

    var myListControl = $("#" + DayControlId);
    myListControl.empty();
    myListControl.append($("<option>").val("").html("D"));

    for (i = 1; i <= daysInMonth; i++) {
        myListControl.append($("<option>").val(i).html(i));
    }

    if (currentDay <= daysInMonth) {
        myListControl.val(currentDay);
    }

    myListControl.selectmenu("refresh");
}

function buildStreetChooserList(listOfPlayers, whichScreenToShow) {
    $.ajax({
        url: '../ajax/Waiver.asp?act=3&pi=' + listOfPlayers,
        error: function () { alert("problems encountered \n Error Code: wnBSC1"); },
        success: function (data) {
            var radhtml = '<div data-role="fieldcontain" style="text-align: center;"><fieldset data-role="controlgroup" style="text-align: center;">';
            radhtml = radhtml + data;
            radhtml = radhtml + '</fieldset></div>';
            switch(whichScreenToShow) {
                case 0:
                    $('#divStreetsRadios').html(radhtml);
                    $("#divStreetsRadios").trigger("create")
                    break;
                case 1: 
                    $('#divWaiverStreetRadios').html(radhtml);
                    $("#divWaiverStreetRadios").trigger("create")
                    break;
            }
        }
    });


}

function checkGroup(selectedValue) {
    if (selectedValue == -1) {
        $("#rowTime").css("display", "table-row");
    } else {
        $("#rowTime").css("display", "none");
    }
}

function checkIn() {
    $("#divStreetChooser").css('display', 'none');
    $('#divEmailMeSearchWaiverNOResults').css('display', 'none');
    $("#divEmailMeSearchWaiver").css('display', 'none');
    $('#divEmailMeSearchWaiverResultsDisplay').html("&nbsp;");
    $('#divEmailMeSearchWaiverResults').css('display', 'none');
    $("#divSearchResults").css("display", "none");
    $("#divNoSearchResults").css("display", "none");
    $("#divSearchNameDOB").css("display", "none");
    $("#divPrompt").css("display", "none");
    $("#divInvalidWaiver").css('display', 'none');
    $("#divMessage1").css('display', 'none');

    $("#divPlayDate").css('display', 'block');
    $("#divPlayTime").css('display', 'block');
    $("#divFinishCheckin").css('display', 'block');

}

function getWaiverIdByPlayer(playerid) {
    var returnValue = "";

    $.ajax({
        url: '../ajax/Waiver.asp?act=7&pd=' + playerid,
        async: false,
        error: function () { alert("problems encountered\nError Code: wnGwBP1"); },
        success: function (data) {
            returnValue = data;
        }
    });

    return returnValue;
}

function getSignatureNameString(FirstName) {
    if (FirstName.length == 0) {
        return "Signature: ";
    } else {
        var name = FirstName;
        if (FirstName.charAt(FirstName.length - 1) == 's') {
            name += '\' signature:';
        } else {
            name += '\'s signature:';
        }

        return name;
    }
}

function initWizard() {
    $("#divStreetChooser").css('display', 'none');
    $('#divEmailMeSearchWaiverNOResults').css('display', 'none');
    $("#divEmailMeSearchWaiver").css('display', 'none');
    $('#divEmailMeSearchWaiverResultsDisplay').html("&nbsp;");
    $('#divEmailMeSearchWaiverResults').css('display', 'none');
    $("#divSearchResults").css("display", "none");
    $("#divNoSearchResults").css("display", "none");
    $("#divSearchNameDOB").css("display", "none");
    $("#divPrompt").css("display", "block");
    $("#divInvalidWaiver").css('display', 'none');
    $("#divMessage1").css('display', 'none');
    $("#divPlayDate").css('display', 'none');
    $("#divPlayTime").css('display', 'none');
    $("#divFinishCheckin").css('display', 'none');


}

function loadCalendar(day) {
    $.ajax({
        url: '../ajax/Calendar.asp?ct=ci&dt=' + day,
        error: function () { alert('problems encountered'); },
        success: function (data) { $('#calendar').html(data); }
    });
}

function loadGroups(selectedDate) {
    $.ajax({
        url: '../ajax/Waiver.asp?act=4&dt=' + selectedDate,
        error: function () { alert("problems encountered"); },
        success: function (data) {
            $("#ddlEventGroup").html(data).selectmenu("refresh");
        }
    });
}

function loadOpenCalendar(day) {
    $.ajax({
        url: '../ajax/Calendar.asp?ct=op&dt=' + day,
        error: function () { alert('problems encountered'); },
        success: function (data) { $('#calendar').html(data); }
    });
}

function loadTimes(selectedDate) {
    $.ajax({
        url: '../ajax/availability.asp?av=2&dt=' + selectedDate,
        error: function () { alert("problems encountered"); },
        success: function (data) {
            $("#ddlPlayTime").html(data).selectmenu("refresh");
        }
    });
}

function navigateBack() {
    window.scrollTo(0, 0);

    switch (wizardStep) {
        case 0: // legalese selection (should never be called)
            $("#divParicipantSelection").css("display", "none");
            $("#divLegalese").css("display", "block");

            wizardStep = 0;
            break;
        case 1: // participant selection (should never be called)
            $("#divParicipantSelection").css("display", "block");
            $("#divSelf").css("display", "none");
            $("#divGuardian").css("display", "none");

            currentMinor = 0;
            wizardStep = 0;
            break;
        case 2: // participant information
            switch (currentMinor) {
                case 0: // gaurdian
                    $("#divParicipantSelection").css("display", "block");
                    $("#divSelf").css("display", "none");
                    $("#divGuardian").css("display", "none");
                    wizardStep = 1;

                    break;
                case 1:
                    $("#divMinor1").css("display", "none");
                    $("#divGuardian").css("display", "block");
                    currentMinor -= 1;
                    break;
                case 2:
                    $("#divMinor2").css("display", "none");
                    $("#divMinor1").css("display", "block");
                    currentMinor -= 1;
                    break;
                case 3:
                    $("#divMinor3").css("display", "none");
                    $("#divMinor2").css("display", "block");
                    currentMinor -= 1;
                    break;
                case 4:
                    $("#divMinor4").css("display", "none");
                    $("#divMinor3").css("display", "block");
                    currentMinor -= 1;
                    break;
                case 5:
                    $("#divMinor5").css("display", "none");
                    $("#divMinor4").css("display", "block");
                    currentMinor -= 1;
                    break;
                case 6:
                    $("#divMinor6").css("display", "none");
                    $("#divMinor5").css("display", "block");
                    currentMinor -= 1;
                    break;
            }
            break;
        case 3: // play date time
            $("#divPlayDate").css("display", "none");

            if (selectedPath == 0) {
                $("#divSelf").css("display", "block");
            } else {
                currentMinor -= 1;
                switch (currentMinor) {
                    case 0: // should not hit this coming back from the signatures
                        if (selectedPath == 1) {
                            $("#divGuardian").css("display", "block");
                        } else if (selectedPath == 2) {
                            $("#divSelf").css("display", "block");
                        }
                        break;
                    case 1:
                        $("#divMinor1").css("display", "block");
                        break;
                    case 2:
                        $("#divMinor2").css("display", "block");
                        break;
                    case 3:
                        $("#divMinor3").css("display", "block");
                        break;
                    case 4:
                        $("#divMinor4").css("display", "block");
                        break;
                    case 5:
                        $("#divMinor5").css("display", "block");
                        break;
                    case 6:
                        $("#divMinor6").css("display", "block");
                        break;
                }
            }

            wizardStep = 2;
            break;
        case 4:
            $("#divPlayTime").css("display", "none");
            $("#divPlayDate").css("display", "block");
            wizardStep = 3;
            break;
        case 5: //signatures
            $("#divSigContainer").css("display", "none");
            $("#divPlayTime").css("display", "block");
            wizardStep = 4;
            break;
        case 6: //complete
            break;
    }
}

function navigateNext() {
    window.scrollTo(0, 0);

    switch (wizardStep) {
        case 0: // legalese
            $("#divLegalese").css("display", "none");
            $("#divParicipantSelection").css("display", "block");

            wizardStep = 1;
            break;
        case 1: // participant selection
            $("#divParicipantSelection").css("display", "none");
            if (selectedPath == 0 || selectedPath == 2) {
                $("#divSelf").css("display", "block");
            }
            else {
                $("#divGuardian").css("display", "block");
            }

            currentMinor = 0;
            wizardStep = 2;
            break;
        case 2: // participant information
            if (selectedPath == 0) {
                //$("#rowParentSignature td:first-child").html(getSignatureNameString($("#txtSelfFirstName").val()));
                $("#divSigName").html(getSignatureNameString($("#txtSelfFirstName").val()));
                $("#divSelf").css("display", "none");
                $("#divPlayDate").css("display", "block");
                wizardStep = 3;
            } else {
                switch (currentMinor) {
                    case 0:
                        if (selectedPath == 1) {
                            $("#divGuardian").css("display", "none");
                            //$("#rowParentSignature td:first-child").html(getSignatureNameString($("#txtParentFirstName").val()));
                            $("#divSigName").html(getSignatureNameString($("#txtParentFirstName").val()));
                        } else if (selectedPath == 2) {
                            $("#divSelf").css("display", "none");
                            //$("#rowParentSignature td:first-child").html(getSignatureNameString($("#txtSelfFirstName").val()));
                            $("#divSigName").html(getSignatureNameString($("#txtSelfFirstName").val()));
                        }
                        $("#divMinor1").css("display", "block");
                        currentMinor += 1;
                        break;
                    case 1:
                        $("#divMinor1").css("display", "none");
                        //$("#rowMinor1Signature").css("display", "table-row");
                        //$("#rowMinor1Signature td:first-child").html(getSignatureNameString($("#txtMinor1FirstName").val()));

                        currentMinor += 1;
                        if (currentMinor > numberOfMinors) {
                            $("#divPlayDate").css("display", "block");
                            wizardStep = 3;
                        } else {
                            $("#divMinor2").css("display", "block");
                        }
                        break;
                    case 2:
                        $("#divMinor2").css("display", "none");
                        //$("#rowMinor2Signature").css("display", "table-row");
                        //$("#rowMinor2Signature td:first-child").html(getSignatureNameString($("#txtMinor2FirstName").val()));

                        currentMinor += 1;
                        if (currentMinor > numberOfMinors) {
                            $("#divPlayDate").css("display", "block");
                            wizardStep = 3;
                        } else {
                            $("#divMinor3").css("display", "block");
                        }
                        break;
                    case 3:
                        $("#divMinor3").css("display", "none");
                        //$("#rowMinor3Signature").css("display", "table-row");
                        //$("#rowMinor3Signature td:first-child").html(getSignatureNameString($("#txtMinor3FirstName").val()));

                        currentMinor += 1;
                        if (currentMinor > numberOfMinors) {
                            $("#divPlayDate").css("display", "block");
                            wizardStep = 3;
                        } else {
                            $("#divMinor4").css("display", "block");
                        }
                        break;
                    case 4:
                        $("#divMinor4").css("display", "none");
                        //$("#rowMinor4Signature").css("display", "table-row");
                        //$("#rowMinor4Signature td:first-child").html(getSignatureNameString($("#txtMinor4FirstName").val()));

                        currentMinor += 1;
                        if (currentMinor > numberOfMinors) {
                            $("#divPlayDate").css("display", "block");
                            wizardStep = 3;
                        } else {
                            $("#divMinor5").css("display", "block");
                        }
                        break;
                    case 5:
                        $("#divMinor5").css("display", "none");
                        //$("#rowMinor5Signature").css("display", "table-row");
                        //$("#rowMinor5Signature td:first-child").html(getSignatureNameString($("#txtMinor5FirstName").val()));

                        currentMinor += 1;
                        if (currentMinor > numberOfMinors) {
                            $("#divPlayDate").css("display", "block");
                            wizardStep = 3;
                        } else {
                            $("#divMinor6").css("display", "block");
                        }
                        break;
                    case 6:
                        $("#divMinor6").css("display", "none");
                        //$("#divPlayDate").css("display", "block");
                        //$("#rowMinor6Signature").css("display", "table-row");
                        $("#rowMinor6Signature td:first-child").html(getSignatureNameString($("#txtMinor6FirstName").val()));
                        wizardStep = 3;
                        currentMinor += 1;
                        break;
                }
            }


            break;
        case 3: // play date
            $("#divPlayDate").css("display", "none");
            $("#divPlayTime").css("display", "block");
            wizardStep = 4;
            break;
        case 4: // play time
            $("#divPlayTime").css("display", "none");
            //$("#divSignatures").css("display", "block");
            $("#divSigContainer").css("display", "block");
            $("#divSig").jSignature();
            wizardStep = 5;
            break;
        case 5: // signatures
            //$("#divSignatures").css("display", "none");
            //$("#divLegalese").css("display", "block");
            //wizardStep = 6;
            break;
        case 6: //complete
            break;
    }
}

function printWaiver() {
    $("#divPromptContainer").css("display", "none");
    $("#printLink").css("display", "none");
    window.print();
    $("#divPromptContainer").css("display", "block");
    $("#printLink").css("display", "block");
}

function saveSignature(HashId) {
    var sigChunks = $("#divSig").jSignature('getData', 'image/svg+xml')[1].chunk(1000);

    saveSignatureRecursive(HashId, sigChunks, 0);
}

function saveSignatureRecursive(HashId, SignatureChunks, CurrentIndex) {
    if (CurrentIndex < SignatureChunks.length) {
        $.ajax({
            url: '../ajax/waiver.asp?act=45&id=' + HashId + '&si=' + $.URLEncode(SignatureChunks[CurrentIndex]),
            error: function (xhr, ajaxOptions, thrownError) {
                if (xhr.status > 0) {
                    //alert("problems encountered sending signature\n\nResponseText = " + xhr.responseText + "\n\nStatus = " + xhr.status + "\n\nThrownError = " + thrownError);
                    alert("problems encountered sending signature");
                }
            },
            success: function () {
                var newIndex = CurrentIndex + 1;
                saveSignatureRecursive(HashId, SignatureChunks, newIndex);
            }
        });
    } else {
        currentMinor += 1;
        if (currentMinor >= numberOfMinors) { document.forms[0].submit(); }
    }
}

function saveWaivers() {
    var playDate = $("#txtPlayDate").val();
    var playTime = $("#ddlPlayTime").val();
    var emailList = ($("#chkEmailList").val() == "on" ? "true" : "false");
    var eventId = "";

    if ($("#txtEventId").val().length > 0) {
        eventId = $("#txtEventId").val();
    } else if ($("#ddlGrouped").val() == "yes") {
        eventId = $("#ddlEventGroup").val();
    }

    currentMinor = 0;

    switch (selectedPath) {
        case 0:
            $.ajax({
                url: '../ajax/waiver.asp?act=40&fn=' + $("#txtSelfFirstName").val() +
                                           '&ln=' + $("#txtSelfLastName").val() +
                                           '&db=' + $("#ddlSelfDOBMonth").val() +
                                           '/' + $("#ddlSelfDOBDay").val() +
                                           '/' + $("#ddlSelfDOBYear").val() +
                                           '&pn=' + $("#txtSelfPhoneNumber").val() +
                                           '&em=' + $("#txtSelfEmailAddress").val() +
                                           '&ad=' + $("#txtSelfAddress").val() +
                                           '&cy=' + $("#txtSelfCity").val() +
                                           '&st=' + $("#ddlSelfState").val() +
                                           '&z=' + $("#txtSelfZipCode").val() +
                                           '&dt=' + playDate +
                                           '&pt=' + playTime +
                                           '&ev=' + eventId +
                                           '&el=' + emailList,
                error: function () { alert("problems encountered saving self"); },
                success: function (data) {
                    $("#txtAdultHash").val(data);
                    saveSignature(data);
                }
            });
            break;
        case 1:
            var parentFirstName = $("#txtParentFirstName").val();
            var parentLastName = $("#txtParentLastName").val();
            var address = $("#txtParentAddress").val();
            var city = $("#txtParentCity").val();
            var state = $("#ddlParentState").val();
            var zip = $("#txtParentZipCode").val();
            var phone = $("#txtParentPhoneNumber").val();
            var email = $("#txtParentEmailAddress").val();


            for (i = 1; i <= numberOfMinors; i++) {
                $.ajax({
                    url: '../ajax/waiver.asp?act=41&fn=' + $("#txtMinor" + i + "FirstName").val() +
                                               '&ln=' + $("#txtMinor" + i + "LastName").val() +
                                               '&pf=' + parentFirstName +
                                               '&pl=' + parentLastName +
                                               '&db=' + $("#ddlMinor" + i + "DOBMonth").val() +
                                               '/' + $("#ddlMinor" + i + "DOBDay").val() +
                                               '/' + $("#ddlMinor" + i + "DOBYear").val() +
                                               '&pn=' + phone +
                                               '&em=' + email +
                                               '&ad=' + address +
                                               '&cy=' + city +
                                               '&st=' + state +
                                               '&z=' + zip +
                                               '&dt=' + playDate +
                                               '&pt=' + playTime +
                                               '&ev=' + eventId +
                                               '&el=' + emailList +
                                               '&id=' + i,
                    error: function () { alert("problems encountered saving minor " + i); },
                    success: function (data) {
                        var returnParts = data.split(",");
                        $("#txtMinor" + returnParts[1] + "Hash").val(returnParts[0]);
                        saveSignature(returnParts[0]);
                    }
                });
            }

            break;
        case 2:
            var parentFirstName = $("#txtSelfFirstName").val();
            var parentLastName = $("#txtSelfLastName").val();
            var address = $("#txtSelfAddress").val();
            var city = $("#txtSelfCity").val();
            var state = $("#ddlSelfState").val();
            var zip = $("#txtSelfZipCode").val();
            var phone = $("#txtSelfPhoneNumber").val();
            var email = $("#txtSelfEmailAddress").val();

            $.ajax({
                url: '../ajax/waiver.asp?act=40&fn=' + parentFirstName +
                                           '&ln=' + parentLastName +
                                           '&db=' + $("#ddlSelfDOBMonth").val() +
                                           '/' + $("#ddlSelfDOBDay").val() +
                                           '/' + $("#ddlSelfDOBYear").val() +
                                           '&pn=' + phone +
                                           '&em=' + email +
                                           '&ad=' + address +
                                           '&cy=' + city +
                                           '&st=' + state +
                                           '&z=' + zip +
                                           '&dt=' + playDate +
                                           '&pt=' + playTime +
                                           '&ev=' + eventId +
                                           '&el=' + emailList,
                error: function () { alert("problems encountered saving self"); },
                success: function (data) {
                    $("#txtAdultHash").val(data);
                    saveSignature(data);
                }
            });

            for (i = 1; i <= numberOfMinors; i++) {
                $.ajax({
                    url: '../ajax/waiver.asp?act=41&fn=' + $("#txtMinor" + i + "FirstName").val() +
                                               '&ln=' + $("#txtMinor" + i + "LastName").val() +
                                               '&pf=' + parentFirstName +
                                               '&pl=' + parentLastName +
                                               '&db=' + $("#ddlMinor" + i + "DOBMonth").val() +
                                               '/' + $("#ddlMinor" + i + "DOBDay").val() +
                                               '/' + $("#ddlMinor" + i + "DOBYear").val() +
                                               '&pn=' + phone +
                                               '&em=' + email +
                                               '&ad=' + address +
                                               '&cy=' + city +
                                               '&st=' + state +
                                               '&z=' + zip +
                                               '&dt=' + playDate +
                                               '&pt=' + playTime +
                                               '&ev=' + eventId +
                                               '&el=' + emailList +
                                               '&id=' + i,
                    error: function () { alert("problems encountered saving minor " + i); },
                    success: function (data) {
                        var returnParts = data.split(",");
                        $("#txtMinor" + returnParts[1] + "Hash").val(returnParts[0]);
                        saveSignature(returnParts[0]);
                    }
                });
            }
            break;
    }

    //window.location.href = "confirm.asp";
    $("#btnFinish").removeAttr("disabled");
}

function setDate(day) {
    $("#txtPlayDate").val(day);
    showCalendar();
    loadGroups(day);
    loadTimes(day);
}

function setNumberOfMinors(MinorCount) {
    $("#cellParticipantNext").css("display", "table-cell");

    numberOfMinors = MinorCount;

    //navigateNext();

    //if (selectedPath == 1)
    //    showParentWithMinorPath();
    //else
    showAdultWithMinorPath();

    $("#txtNumberOfMinors").val(MinorCount);
}

function setupWizardPath(PathId) {
    switch (PathId) {
        case 0: // adult only
            $("#rowMinors").css("display", "none");
            $("#cellParticipantNext").css("display", "table-cell");
            selectedPath = PathId;
            //navigateNext();
            showAdultPath();
            break;
        case 1: // minors only
        case 2: // adult and minros
            $("#rowMinors").css("display", "table-row");
            $("#cellParticipantNext").css("display", "none");
            selectedPath = PathId;
            break;
    }
    $("#txtSelectedPath").val(PathId);
}

function showAdultPath() {
    currentMinor = 0;
    // show these panels
    $("#divSelf").css("display", "block");

    var path = window.location.pathname.toLowerCase();
    var pageName = path.substring(path.lastIndexOf('/') + 1);

    if ($("#txtUseRegistration").val().toLowerCase() == "true" && $("#txtEventId").val().length == 0) {
        $("#divPlayDate").css("display", "block");
        $("#divPlayTime").css("display", "block");
    } else if ($("#txtUseRegistration").val().toLowerCase() == "false" && (pageName === "default.asp" || pageName === "")) {
        $("#divPlayDate").css("display", "block");
        $("#divPlayTime").css("display", "block");
    }

    if ($("#divSigContainer").css("display") == "none") {
        $("#divSigContainer").css("display", "block");
        $("#divSig").jSignature();
    }
    $("#divWaiverCustomField").css("display", "block");

    // hide these panels
    $("#divGuardian").css("display", "none");
    $("#divMinor1").css("display", "none");
    $("#divMinor2").css("display", "none");
    $("#divMinor3").css("display", "none");
    $("#divMinor4").css("display", "none");
    $("#divMinor5").css("display", "none");
    $("#divMinor6").css("display", "none");
}

function showAdultWithMinorPath() {
    currentMinor = 0;
    // show these panels
    $("#divSelf").css("display", "block");

    var path = window.location.pathname.toLowerCase();
    var pageName = path.substring(path.lastIndexOf('/') + 1);

    if ($("#txtUseRegistration").val().toLowerCase() == "true" && $("#txtEventId").val().length == 0) {
        $("#divPlayDate").css("display", "block");
        $("#divPlayTime").css("display", "block");
    } else if ($("#txtUseRegistration").val().toLowerCase() == "false" && (pageName === "default.asp" || pageName === "")) {
        $("#divPlayDate").css("display", "block");
        $("#divPlayTime").css("display", "block");
    }

    if ($("#divSigContainer").css("display") == "none") {
        $("#divSigContainer").css("display", "block");
        $("#divSig").jSignature();
    }

    $("#divWaiverCustomField").css("display", "block");

    for (i = 0; i <= 6; i++) {
        if (i <= numberOfMinors)
            $("#divMinor" + i).css("display", "block");
        else
            $("#divMinor" + i).css("display", "none");
    }

    // hide these panels
    $("#divGuardian").css("display", "none");
}

function showCalendar() {
    var row = $("#rowCalendar");
    var month = $("#rowMonth");

    if (row.css("display") == "none") {
        row.css("display", "table-row");
        month.css("display", "table-row");
    } else {
        row.css("display", "none");
        month.css("display", "none");
    }
}

function showHideWaiverForm(which) {
    switch (which) {
        case 0: //hide
            $("#divWaiver").css("display", "none");
            break;
        case 1: //show
            $("#divWaiver").css("display", "block");
            break;
        default:
            startNewWaiver();
    }
}

function showParentWithMinorPath() {
    currentMinor = 0;
    // show these panels
    $("#divGuardian").css("display", "block");

    var path = window.location.pathname.toLowerCase();
    var pageName = path.substring(path.lastIndexOf('/') + 1);

    if ($("#txtUseRegistration").val().toLowerCase() == "true" && $("#txtEventId").val().length == 0) {
        $("#divPlayDate").css("display", "block");
        $("#divPlayTime").css("display", "block");
    } else if ($("#txtUseRegistration").val().toLowerCase() == "false" && (pageName === "default.asp" || pageName === "")) {
        $("#divPlayDate").css("display", "block");
        $("#divPlayTime").css("display", "block");
    }

    if ($("#divSigContainer").css("display") == "none") {
        $("#divSigContainer").css("display", "block");
        $("#divSig").jSignature();
    }

    for (i = 0; i <= 6; i++) {
        if (i <= numberOfMinors)
            $("#divMinor" + i).css("display", "block");
        else
            $("#divMinor" + i).css("display", "none");
    }

    $("#divWaiverCustomField").css("display", "block");

    // hide these panels
    $("#divSelf").css("display", "none");
}

function showNavigationButtons(ShowBack, ShowNext) {
    if (ShowBack || ShowNext) {
        $("#divNavigationButtons").css("display", "block");

        if (ShowBack)
            $("#btnBack").css("display", "inline");
        else
            $("#btnBack").css("display", "none");

        if (ShowNext)
            $("#btnNext").css("display", "inline");
        else
            $("#btnNext").css("display", "none");

    } else {
        $("#divNavigationButtons").css("display", "none");
    }
}

function startNewWaiver() {
    $("#divStreetChooser").css('display', 'none');
    $('#divEmailMeSearchWaiverNOResults').css('display', 'none');
    $("#divEmailMeSearchWaiver").css('display', 'none');
    $('#divEmailMeSearchWaiverResultsDisplay').html("&nbsp;");
    $('#divEmailMeSearchWaiverResults').css('display', 'none');
    $("#divSearchResults").css("display", "none");
    $("#divNoSearchResults").css("display", "none");
    $("#divSearchNameDOB").css("display", "none");
    $("#divPrompt").css("display", "none");
    $("#divLegalese").css("display", "block");
    $("#divInvalidWaiver").css('display', 'none');
    $("#divMessage1").css('display', 'none');
    $("#divWaiver").css('display', 'block');
    wizardStep = 0;

    $('input[name=txtSelfPid]').val($('button[name=btnNewExpiredWaiver]').val());
}

function startNewWaiverFromSearch() {
    var age = getAge($("#ddlSearchDOBMonth").val() + "/" + $("#ddlSearchDOBDay").val() + "/" + $("#ddlSearchDOBYear").val());

    if (age <= 17) {
        $("#txtMinor1FirstName").val($("#txtSearchFirstName").val());
        $("#txtMinor1LastName").val($("#txtSearchLastName").val());
    } else {
        $("#txtSelfFirstName").val($("#txtSearchFirstName").val());
        $("#txtSelfLastName").val($("#txtSearchLastName").val());
    }
    
    startNewWaiver();
}

function submitForm() {
    $("#txtSignature").val($("#divSig").jSignature('getData', 'image/svg+xml')[1]);
    if (validateWaiverForm()) {
        this.disabled = true;
        //saveWaivers();
        return true;
    } else {
        return false;
    }

}

function submitCheckIn() {
    if (validateCheckIn()) {
        this.disabled = true;
        $("#txtCheckingIn").val("true");
        $("#txtCheckInPID").val(checkInId);
        $("#myForm").submit();
    } else {
        return false;
    }
}

    function toggleGroupTime(value) {
        if (value == "yes") {
            $("#rowGroup").css("display", "table-row");
            $("#rowTime").css("display", "none");
        } else if (value == "no") {
            $("#rowGroup").css("display", "none");
            $("#rowTime").css("display", "table-row");
        } else {
            $("#rowGroup").css("display", "none");
            $("#rowTime").css("display", "none");
        }
    }