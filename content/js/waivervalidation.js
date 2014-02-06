function validateConfirmEmail(emailTextFieldId, confirmTextFieldId, matchValidatorFieldId) {
    var valid = true;
    var confirmAddress = $("#" + confirmTextFieldId);
    var emailAddress = $("#" + emailTextFieldId);

    if (emailAddress.val() != confirmAddress.val()) {
        valid = false;
        $("#" + matchValidatorFieldId).css("display", "block");
        confirmAddress.addClass("error");
    } else {
        $("#" + matchValidatorFieldId).css("display", "none");
        confirmAddress.removeClass("error");
    }

    return valid;
}

function validateRequiredCustomElements(location) {
    var valid = true;
    var values = jQuery("#myForm").serializeArray();
    /* Because serializeArray() ignores unset checkboxes and radio buttons: */
    values = values.concat(
            jQuery('#myForm input[type=checkbox]:not(:checked)').map(
                    function () {
                        return { "name": this.name, "value": this.value }
                    }).get()
    );

    for (var i=0;i<values.length;i++){
        if (values[i].name.substr(0, 6) == "Custom") {
            if (values[i].name.substr(values[i].name.indexOf('_') + 1, location.length) == location && values[i].name.substr(values[i].name.lastIndexOf('_') + 1, 3) == 'req') {
                switch (document.getElementById(values[i].name).type.toUpperCase()) {
                    case 'SELECT':
                        valid = validateRequiredDDL(values[i].name, 'val' + values[i].name);
                        break;
                    case 'CHECKBOX':
                        valid = validateRequiredCheckBox(values[i].name, 'val' + values[i].name);
                        break;
                    default:
                        valid = validateRequired(values[i].name, 'val' + values[i].name);
                }
            }
        }
    }

    return valid;
}

function validateDate(monthFieldId, dayFieldId, yearFieldId) {
    var ddlMonth = $("#" + monthFieldId);
    var ddlDay = $("#" + dayFieldId);
    var ddlYear = $("#" + yearFieldId);

    if (ddlYear.val() == "") {
        ddlYear.parent().addClass("error");
    } else {
        ddlYear.parent().removeClass("error");
    }

    if (ddlMonth.val() == "") {
        ddlMonth.parent().addClass("error");
    } else {
        ddlMonth.parent().removeClass("error");
    }

    if (ddlDay.val() == "") {
        ddlDay.parent().addClass("error");
    } else {
        ddlDay.parent().removeClass("error");
    }
}

function validateDOB(monthFieldId, dayFieldId, yearFieldId, requiredValidationFieldId) {
    var valid = true;
    var ddlMonth = $("#" + monthFieldId);
    var ddlDay = $("#" + dayFieldId);
    var ddlYear = $("#" + yearFieldId);

    if (ddlMonth.val() == "" || ddlDay.val() == "" || ddlYear.val() == "") {
        valid = false;
        $("#" + requiredValidationFieldId).css("display", "block");
        ddlMonth.parent().addClass("error");
        ddlDay.parent().addClass("error");
        ddlYear.parent().addClass("error");
    } else {
        $("#" + requiredValidationFieldId).css("display", "none");
        ddlMonth.parent().removeClass("error");
        ddlDay.parent().removeClass("error");
        ddlYear.parent().removeClass("error");
    }

    return valid;
}

function validateDOBForAdult(monthFieldId, dayFieldId, yearFieldId, verificationValidationFieldId) {
    var age = 0;
    var valid = true;

    if ($("#" + monthFieldId).val().length == 0 || $("#" + dayFieldId).val().length == 0 || $("#" + yearFieldId).val().length == 0) {
        age = 0;
    } else {
        /*
        try {
            var today = new Date();
            var dob = new Date($("#" + monthFieldId).val() + "/" + $("#" + dayFieldId).val() + "/" + $("#" + yearFieldId).val());

            var todayInMS = today.getTime();
            var dobInMS = dob.getTime();

            var diffInMS = Math.abs(todayInMS - dobInMS);

            // format is milliseconds, seconds, minutes, hours, days
            age = Math.floor(diffInMS / (1000 * 60 * 60 * 24 * 365));
        }
        catch (err) {
            age = 0;
        }
        */
        var age = getAge($("#" + monthFieldId).val() + "/" + $("#" + dayFieldId).val() + "/" + $("#" + yearFieldId).val());
    }

    if (age < 18) {
        $("#" + verificationValidationFieldId).css("display", "block");
        $("#" + monthFieldId).parent().addClass("error");
        $("#" + dayFieldId).parent().addClass("error");
        $("#" + yearFieldId).parent().addClass("error");
        valid = false;
    } else {
        $("#" + verificationValidationFieldId).css("display", "none");
        $("#" + monthFieldId).parent().removeClass("error");
        $("#" + dayFieldId).parent().removeClass("error");
        $("#" + yearFieldId).parent().removeClass("error");
    }

    return valid;
}

function validateDOBForMinimumAge(minimumAge, monthFieldId, dayFieldId, yearFieldId, verificationValidationFieldId) {
    var age = 0;
    var valid = true;

    if ($("#" + monthFieldId).val().length == 0 || $("#" + dayFieldId).val().length == 0 || $("#" + yearFieldId).val().length == 0) {
        age = 0;
    } else {
        /*
        try {
            var today = new Date();
            var dob = new Date($("#" + monthFieldId).val() + "/" + $("#" + dayFieldId).val() + "/" + $("#" + yearFieldId).val());

            var todayInMS = today.getTime();
            var dobInMS = dob.getTime();

            var diffInMS = Math.abs(todayInMS - dobInMS);

            // format is milliseconds, seconds, minutes, hours, days
            age = Math.floor(diffInMS / (1000 * 60 * 60 * 24 * 365));
        }
        catch (err) {
            age = 0;
        }
        */
        var age = getAge($("#" + monthFieldId).val() + "/" + $("#" + dayFieldId).val() + "/" + $("#" + yearFieldId).val());
    }

    if (age < minimumAge) {
        $("#" + verificationValidationFieldId).css("display", "block");
        $("#" + monthFieldId).parent().addClass("error");
        $("#" + dayFieldId).parent().addClass("error");
        $("#" + yearFieldId).parent().addClass("error");
        valid = false;
    } else {
        $("#" + verificationValidationFieldId).css("display", "none");
        $("#" + monthFieldId).parent().removeClass("error");
        $("#" + dayFieldId).parent().removeClass("error");
        $("#" + yearFieldId).parent().removeClass("error");
    }

    return valid;
}

function validateDOBForMinor(monthFieldId, dayFieldId, yearFieldId, verificationValidationFieldId) {
    var age = 0;
    var valid = true;

    if ($("#" + monthFieldId).val().length == 0 || $("#" + dayFieldId).val().length == 0 || $("#" + yearFieldId).val().length == 0) {
        age = 0;
    } else {
        /*
        try {
            var today = new Date();
            var dob = new Date($("#" + monthFieldId).val() + "/" + $("#" + dayFieldId).val() + "/" + $("#" + yearFieldId).val());

            var todayInMS = today.getTime();
            var dobInMS = dob.getTime();

            var diffInMS = Math.abs(todayInMS - dobInMS);

            // format is milliseconds, seconds, minutes, hours, days
            age = Math.floor(diffInMS / (1000 * 60 * 60 * 24 * 365));
        }
        catch (err) {
            age = 0;
        }
        */
        var age = getAge($("#" + monthFieldId).val() + "/" + $("#" + dayFieldId).val() + "/" + $("#" + yearFieldId).val());
    }

    if (age >= 18) {
        $("#" + verificationValidationFieldId).css("display", "block");
        $("#" + monthFieldId).parent().addClass("error");
        $("#" + dayFieldId).parent().addClass("error");
        $("#" + yearFieldId).parent().addClass("error");
        valid = false;
    } else {
        $("#" + verificationValidationFieldId).css("display", "none");
        $("#" + monthFieldId).parent().removeClass("error");
        $("#" + dayFieldId).parent().removeClass("error");
        $("#" + yearFieldId).parent().removeClass("error");
    }

    return valid;
}

function validateEmail(textFieldId, requiredValidationFieldId, formatValidationFieldId) {
    var valid = true;
    var emailRegExp = /^([a-zA-Z0-9_\.\-\+])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;
    var email = $("#" + textFieldId);

    if (email.val().length == 0) {
        valid = false;
        $("#" + requiredValidationFieldId).css("display", "block");
        $("#" + formatValidationFieldId).css("display", "none");
        email.addClass("error");
    } else if (emailRegExp.test(email.val()) == false) {
        valid = false;
        $("#" + formatValidationFieldId).css("display", "block");
        $("#" + requiredValidationFieldId).css("display", "none");
        email.addClass("error");
    } else {
        $("#" + requiredValidationFieldId).css("display", "none");
        $("#" + formatValidationFieldId).css("display", "none");
        email.removeClass("error");
    }
    return valid;
}

function validateFormMinor1() {
    var valid;

    valid = !!(validateRequired('txtMinor1FirstName', 'valMinor1FirstNameReq') &
               validateRequired('txtMinor1LastName', 'valMinor1LastNameReq') &
               validateMinorDOB('ddlMinor1DOBMonth', 'ddlMinor1DOBDay', 'ddlMinor1DOBYear', 'valMinor1DOBMinimum', 'valMinor1DOBVerification', 'valMinor1DOBReq') & 
               validateRequiredCustomElements('Minor1') 
               );

    return valid;
}

function validateFormMinor2() {
    var valid;

    valid = !!(validateRequired('txtMinor2FirstName', 'valMinor2FirstNameReq') &
               validateRequired('txtMinor2LastName', 'valMinor2LastNameReq') &
               validateMinorDOB('ddlMinor2DOBMonth', 'ddlMinor2DOBDay', 'ddlMinor2DOBYear', 'valMinor2DOBMinimum', 'valMinor2DOBVerification', 'valMinor2DOBReq') & 
               validateRequiredCustomElements('Minor2')
               );

    return valid;
}

function validateFormMinor3() {
    var valid;

    valid = !!(validateRequired('txtMinor3FirstName', 'valMinor3FirstNameReq') &
               validateRequired('txtMinor3LastName', 'valMinor3LastNameReq') &
               validateMinorDOB('ddlMinor3DOBMonth', 'ddlMinor3DOBDay', 'ddlMinor3DOBYear', 'valMinor3DOBMinimum', 'valMinor3DOBVerification', 'valMinor3DOBReq') & 
               validateRequiredCustomElements('Minor3')
               );

    return valid;
}

 function validateFormMinor4() {
    var valid;

    valid = !!(validateRequired('txtMinor4FirstName', 'valMinor4FirstNameReq') &
               validateRequired('txtMinor4LastName', 'valMinor4LastNameReq') &
               validateMinorDOB('ddlMinor4DOBMonth', 'ddlMinor4DOBDay', 'ddlMinor4DOBYear', 'valMinor4DOBMinimum', 'valMinor4DOBVerification', 'valMinor4DOBReq') & 
               validateRequiredCustomElements('Minor4')
               );

    return valid;
}

function validateFormMinor5() {
    var valid;

    valid = !!(validateRequired('txtMinor5FirstName', 'valMinor5FirstNameReq') &
               validateRequired('txtMinor5LastName', 'valMinor5LastNameReq') &
               validateMinorDOB('ddlMinor5DOBMonth', 'ddlMinor5DOBDay', 'ddlMinor5DOBYear', 'valMinor5DOBMinimum', 'valMinor5DOBVerification', 'valMinor5DOBReq') & 
               validateRequiredCustomElements('Minor5')
               );

    return valid;
}

function validateFormMinor6() {
    var valid;

    valid = !!(validateRequired('txtMinor6FirstName', 'valMinor6FirstNameReq') &
               validateRequired('txtMinor6LastName', 'valMinor6LastNameReq') &
               validateMinorDOB('ddlMinor6DOBMonth', 'ddlMinor6DOBDay', 'ddlMinor6DOBYear', 'valMinor6DOBMinimum', 'valMinor6DOBVerification', 'valMinor6DOBReq') & 
               validateRequiredCustomElements('Minor6')
               );

    return valid;
}

function validateFormParent() {
    var valid;

    valid = !!(validateRequired('txtParentFirstName', 'valParentFirstNameReq') &
               validateRequired('txtParentLastName', 'valParentLastNameReq') &
               validatePhoneNumber('txtParentPhoneNumber', 'valParentPhoneNumberReq', 'valParentPhoneNumberFormat') &
               validateEmail('txtParentEmailAddress', 'valParentEmailAddressReq', 'valParentEmailAddressFormat') &
               validateConfirmEmail('txtParentEmailAddress', 'txtParentConfirmEmail', 'valParentConfirmEmailMatch') &
               validateRequired('txtParentAddress', 'valParentAddressReq') &
               validateRequired('txtParentCity', 'valParentCityReq') &
               validateRequired('txtParentZipCode', 'valParentZipCodeReq')
               );

    return valid;
}

function validateFormPlayDate() {
    var valid = true;

    $("#valPlayDateRequired").css("display", "none");
    $("#txtPlayDate").removeClass("error");

    if ($("#txtEventId").val().length == 0) {
        if ($("#txtPlayDate").val().length == 0) {
            valid = false;
            $("#valPlayDateRequired").css("display", "inline");
            $("#txtPlayDate").addClass("error");
        }
    }

    return valid;
}

function validateFormPlayTime() {
    var valid = true;

    if ($("#txtUseRegistration").val() == "True") {
        $("#valGroupedRequired").css("display", "none");
        $("#ddlGrouped").parent().removeClass("error");
        $("#valEventGroupRequired").css("display", "none");
        $("#ddlEventGroup").parent().removeClass("error");
        $("#valPlayTimeRequired").css("display", "none");
        $("#ddlPlayTime").parent().removeClass("error");

        if ($("#txtEventId").val().length == 0) {
            if ($("#ddlGrouped").val().length == 0) {
                valid = false;
                $("#valGroupedRequired").css("display", "inline");
                $("#ddlGrouped").parent().addClass("error");
            } else {

                if ($("#ddlGrouped").val() == "yes") {
                    if ($("#ddlEventGroup option").length == 0 ||$("#ddlEventGroup").val().length == 0) {
                        valid = false;
                        $("#valEventGroupRequired").css("display", "inline");
                        $("#ddlEventGroup").parent().addClass("error");
                    }
                } else if ($("#ddlGrouped").val() == "no") {
                    if ($("#ddlPlayTime option").length == 0 || $("#ddlPlayTime").val().length == 0) {
                        valid = false;
                        $("#valPlayTimeRequired").css("display", "inline");
                        $("#ddlPlayTime").parent().addClass("error");
                    }
                }
            }
        }
    } else {
        $("#valPlayTimeRequired").css("display", "none");
        $("#ddlPlayTime").parent().removeClass("error");

        if ($("#ddlPlayTime option").length == 0 || $("#ddlPlayTime").val().length == 0) {
            valid = false;
            $("#valPlayTimeRequired").css("display", "inline");
            $("#ddlPlayTime").parent().addClass("error");
        }
    }

    return valid;
}

function validateFormSelf() {
    var valid;

    valid = !!(validateRequired('txtSelfFirstName', 'valSelfFirstNameReq') &
               validateRequiredCustomElements('Self') &
               validateRequired('txtSelfLastName', 'valSelfLastNameReq') &
               validateDOB('ddlSelfDOBMonth', 'ddlSelfDOBDay', 'ddlSelfDOBYear', 'valSelfDOBReq') &
               validateDOBForAdult('ddlSelfDOBMonth', 'ddlSelfDOBDay', 'ddlSelfDOBYear', 'valSelfDBOVerification') &
               validatePhoneNumber('txtSelfPhoneNumber', 'valSelfPhoneNumberReq', 'valSelfPhoneNumberFormat') &
               validateEmail('txtSelfEmailAddress', 'valSelfEmailAddressReq', 'valSelfEmailAddressFormat') &
               //validateConfirmEmail('txtSelfEmailAddress', 'txtSelfConfirmEmail', 'valSelfConfirmEmailMatch') &
               validateRequired('txtSelfAddress', 'valSelfAddressReq') &
               validateRequired('txtSelfCity', 'valSelfCityReq') &
               validateRequired('txtSelfZipCode', 'valSelfZipCodeReq')
               );

    return valid;
}

function validateFormSignature() {
    var valid = true;

    if ($("#divSig").jSignature("getData", "image/svg+xml")[1].length <= 1000) {
        valid = false;
        $("#valSignatureRequired").css("display", "inline");
        $("#divSig").addClass("error");
    } else {
        $("#valSignatureRequired").css("display", "none");
        $("#divSig").removeClass("error");
    }

    return valid;
}

function validateMinorDOB(monthFieldId, dayFieldId, yearFieldId, validationMinimumFieldId, validationVerificationFieldId, valReq) {
    var valid;

    if (validateDOB(monthFieldId, dayFieldId, yearFieldId, valReq)) {
        if (validateDOBForMinimumAge($("#txtMinAge").val(), monthFieldId, dayFieldId, yearFieldId, validationMinimumFieldId)) {
            if (validateDOBForMinor(monthFieldId, dayFieldId, yearFieldId, validationVerificationFieldId)) {
                valid = true;
            } else { valid = false; }
        } else { valid = false; }
    } else { valid = false; }

    return valid;
}

function validatePhoneNumber(textFieldId, requiredValidationFieldId, formatValidationFieldId) {
    var valid = true;
    var phoneRegExp = /^\(?(\d{3})\)?[- ]?(\d{3})[- ]?(\d{4})$/;
    var phoneNumber = $("#" + textFieldId);
    
    if (phoneNumber.val().length == 0) {
        valid = false;
        $("#" + requiredValidationFieldId).css("display", "block");
        $("#" + formatValidationFieldId).css("display", "none");
        phoneNumber.addClass("error");
    } else if (phoneRegExp.test(phoneNumber.val()) == false) {
        valid = false;
        $("#" + formatValidationFieldId).css("display", "block");
        $("#" + requiredValidationFieldId).css("display", "none");
        phoneNumber.addClass("error");
    } else {
        $("#" + requiredValidationFieldId).css("display", "none");
        $("#" + formatValidationFieldId).css("display", "none");
        phoneNumber.removeClass("error");
    }

    return valid;
}

function validateRequired(textFieldId, requiredValidationFieldId) {
    var valid = true;
    var nameTextField = $("#" + textFieldId);

    if (nameTextField.val().length == 0) {
        valid = false;
        $("#" + requiredValidationFieldId).css("display", "block");
        nameTextField.addClass("error");
    } else {
        $("#" + requiredValidationFieldId).css("display", "none");
        nameTextField.removeClass("error");
    }

    return valid;
}

function validateRequiredCheckBox(fieldId, requiredField) {
    var valid = true;
    if (!$("#" + fieldId).is(':checked')) {
        // not checked
        valid = false;
        $("#" + requiredField).css("display", "block");
        $("#" + fieldId).addClass("error");
    } else {
        valid = false;
        $("#" + requiredField).css("display", "none");
        $("#" + fieldId).removeClass("error");
    }
    return valid;
}

function validateRequiredDDL(fieldId, requiredField) {
    var valid = true;
    var nameDDLField = $("#" + fieldId);

    if (nameDDLField.val() == "") {
        valid = false;
        $("#" + requiredField).css("display", "block");
        nameDDLField.parent().addClass("error");
    } else {
        valid = true;
        $("#" + requiredField).css("display", "none");
        nameDDLField.parent().removeClass("error");
    }

    return valid;
}

function validateSearchWaiver(whichSearch) {
    var valid = true;

    //alert("searching by Name and DOB");
    valid = (validateRequired("txtSearchFirstName", "valSearchFirstNameReq") && validateRequired("txtSearchLastName", "valSearchLastNameReq") && validateDOB("ddlSearchDOBMonth", "ddlSearchDOBDay", "ddlSearchDOBYear", "valSearchDOBReq"));

    if (!valid) {
        var errFields = $(".error");
        errFields[0].focus();

        $("#valSearchResults").css("display", "inline");
    } else {
        $("#valSearchResults").css("display", "none");
    }

    return valid;
}

    function validateCheckIn() {
        var valid = true;
        valid = validateFormPlayDate() & validateFormPlayTime();

        if (valid) {
            $.ajax({
                url: '../ajax/Waiver.asp?act=65&pd=' + checkInId + "&pdt=" + $("#txtPlayDate").val(),
                error: function () { alert("problems encountered \n Error Code: wnVCHKIn1"); },
                success: function (data) {
                    if (data.substr(0, 13) != "[NOPLAYDATES]") {
                        alert("This player has already checked in for this date.");
                        valid = false;
                    }
                },
                async: false
            });
        }
        return valid;
    }

    function validateWaiverForm() {
        var valid = true;

        switch (selectedPath) {
            case 0: //adult only
                valid = validateFormSelf() & validateFormSignature() & validateFormPlayDate() & validateFormPlayTime();
                break;
            case 1: //minors only
                //valid = validateFormParent() & validateFormSignature() & validateFormPlayDate() & validateFormPlayTime();
                valid = validateFormSelf() & validateFormSignature() & validateFormPlayDate() & validateFormPlayTime();

                for (i = 1; i <= numberOfMinors; i++) {
                    if (!eval("validateFormMinor" + i + "()"))
                        valid = false;
                }

                break;
            case 2: //adult and minors
                valid = validateFormSelf() & validateFormSignature() & validateFormPlayDate() & validateFormPlayTime();

                for (i = 1; i <= numberOfMinors; i++) {
                    if (!eval("validateFormMinor" + i + "()"))
                        valid = false;
                }

                break;
        }

        var customGeneral = true;
        customGeneral = validateRequiredCustomElements('General');

        if (valid && customGeneral) {
            valid = true;
        } else { valid = false; }
    
        if (!valid) {
            var errFields = $(".error");
            errFields[0].focus();

            $("#valFormResults").css("display", "inline");
        } else {
            $("#valFormResults").css("display", "none");
        }

        return valid;
    }