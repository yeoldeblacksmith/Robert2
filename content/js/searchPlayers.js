

function emailMeSearchWaiver() {
    $('#divEmailMeSearchWaiverNOResults').css('display', 'none');
    $("#divEmailMeSearchWaiver").css('display', 'none');
    $('#divEmailMeSearchWaiverResultsDisplay').html("&nbsp;");
    $('#divEmailMeSearchWaiverResults').css('display', 'none');
    $("#divNoSearchResults").css("display", "none");
    $("#divEmailMeSearchWaiver").css("display", "block");
    $("#divInvalidWaiver").css('display', 'none');
}

function sendWaiverSearchEmail() {
    var validEntry = validateEmail('txtWaiverSearchEmailAddress', 'valWaiverSearchEmailAddressReq', 'valWaiverSearchEmailAddressFormat');
    var submittedEmail = $("#txtWaiverSearchEmailAddress").val();

    if (validEntry) {
        $.ajax({
            url: '../ajax/Waiver.asp?act=2&em=' + $("#txtWaiverSearchEmailAddress").val(),
            error: function (jqXHR, exception) {
                alert('problems encountered');
            },
            success: function (data) {
                if (data != "no email found for waivers") {
                    $("#divEmailMeSearchWaiver").css('display', 'none');
                    $('#divEmailMeSearchWaiverResultsDisplay').html(data);
                    $('#divEmailMeSearchWaiverResults').css('display', 'block');
                } else {
                    $("#divEmailMeSearchWaiver").css('display', 'none');
                    $('#divEmailMeSearchWaiverResults').css('display', 'none');
                    $('#divEmailMeSearchWaiverNOResults').css('display', 'block');
                }

            }
        });
    }
}

function findWaiver(whichSearch) {
    switch (whichSearch) {
        case 0:
            if (validateSearchWaiver(whichSearch)) {
                $("#divStreetChooser").css('display', 'none');
                $('#divEmailMeSearchWaiverNOResults').css('display', 'none');
                $("#divEmailMeSearchWaiver").css('display', 'none');
                $('#divEmailMeSearchWaiverResultsDisplay').html("&nbsp;");
                $('#divEmailMeSearchWaiverResults').css('display', 'none');
                $("#divInvalidWaiver").css('display', 'none');

                showSearchResults(searchForPlayersByNameDOB($("#txtSearchFirstName").val(), $("#txtSearchLastName").val(), $("#ddlSearchDOBYear").val() + '-' + $("#ddlSearchDOBMonth").val() + '-' + $("#ddlSearchDOBDay").val()));
            }
            break;
        case 1:
            alert("problems encountered\nError Code: spFWs1");
            //$("#divWaiverStreet").css('display', 'block');
            //$("#divWaiverStreet").addClass("error");
            //selectedSearch = 1;
            //buildStreetChooserList(searchResults.substring(1), selectedSearch);

            break;
    }

    return false;
}

function searchForPlayersByNameDOB(fn, ln, dob) {
    var returnValue = "";

    $.ajax({
        url: '../ajax/Waiver.asp?act=1&fn=' + fn + '&ln=' + ln + '&db=' + dob,
        async: false,
        error: function () {
            alert("problems encountered\nError Code: spSBNDob1");
            returnValue = "error";
        },
        success: function (data) {
            returnValue = data;
        }
    });

    return returnValue;
}

function searchForPlayersByNameDOBAddress(fn, ln, dob, addy) {
    var returnValue = "";

    $.ajax({
        url: '../ajax/Waiver.asp?act=6&fn=' + fn + '&ln=' + ln + '&db=' + dob + '&ad=' + addy,
        async: false,
        error: function () {
            alert("problems encountered\nError Code: spNDA2");
            returnValue = "error";
        },
        success: function (data) {
            returnValue = data;
        }
    });

    return returnValue;
}

function searchForPlayerById(pId) {
    var returnValue = "";

    if (pId == 'street not listed') {
        returnValue = pId;
    } else {
        $.ajax({
            url: '../ajax/Waiver.asp?act=5&pd=' + pId,
            async: false,
            error: function (jqXHR, exception) {
                //alert("There was a problem encountered while executing this search.")
                if (jqXHR.status === 0) {
                    alert('Not connect.\n Verify Network.');
                } else if (jqXHR.status == 404) {
                    alert('Requested page not found. [404]');
                } else if (jqXHR.status == 500) {
                    alert('Internal Server Error [500].');
                } else if (exception === 'parsererror') {
                    alert('Requested JSON parse failed.');
                } else if (exception === 'timeout') {
                    alert('Time out error.');
                } else if (exception === 'abort') {
                    alert('Ajax request aborted.');
                } else {
                    alert('Uncaught Error.\n' + jqXHR.responseText);
                }
                returnValue = "error";
            },
            success: function (data) {
                returnValue = data;
            }
        });
    }
      
        return returnValue;
}

function showWaiverSearch() {
    $('#divEmailMeSearchWaiverNOResults').css('display', 'none');
    $("#divEmailMeSearchWaiver").css('display', 'none');
    $('#divEmailMeSearchWaiverResultsDisplay').html("&nbsp;");
    $('#divEmailMeSearchWaiverResults').css('display', 'none');
    $("#divSearchResults").css("display", "none");
    $("#divNoSearchResults").css("display", "none");
    $("#divPrompt").css("display", "none");
    $("#divSearchNameDOB").css("display", "block");

}

function showSearchResults(searchResults) {
        switch (parseInt(searchResults.substring(0, 1))) {
            case 1:
                //found 1 player with valid waiver
                checkInId = searchResults.substring(1);
                $("#divSearchResults").css('display', 'block');
                $('#divNoSearchResults').css('display', 'none');
                $("#divStreetChooser").css('display', 'none');
                $("#divInvalidWaiver").css('display', 'none');
                break;
            case 2:
                //found multiple players for search
                selectedSearch = 0;
                buildStreetChooserList(searchResults.substring(1), selectedSearch);
                $("#divNoSearchResults").css('display', 'none');
                $('#divSearchResults').css('display', 'none');
                $("#divStreetChooser").css('display', 'block');
                $("#divInvalidWaiver").css('display', 'none');
                break;
            case 3:
                //found 1 player with  invalid waiver
                $("#divInvalidWaiver").css('display', 'block');
                $('button[name=btnNewExpiredWaiver]').val(searchResults.substring(1));
                break;
            default:
                //no player or valid waiver found
                $("#divSearchResults").css('display', 'none');
                $('#divNoSearchResults').css('display', 'block');
                $("#divStreetChooser").css('display', 'none');
                $("#divInvalidWaiver").css('display', 'none');
                $("#txtSearchOverride").val('true');
        }
}

function streetChosen(whichStreetScreen) {
    switch (whichStreetScreen) {
        case 0:
            $("#divStreetChooser").css('display', 'none');
            $('#divEmailMeSearchWaiverNOResults').css('display', 'none');
            $("#divEmailMeSearchWaiver").css('display', 'none');
            $('#divEmailMeSearchWaiverResultsDisplay').html("&nbsp;");
            $('#divEmailMeSearchWaiverResults').css('display', 'none');

            showSearchResults(searchForPlayerById($('input[name=street-radio]:radio:checked').val()));
            break;
        case 1:
            alert("waiver street");
            break;
        default:
            alert("problem encountered\n Error Code: spSTSe1");
    }
    
    
    
}






