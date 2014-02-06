function getAge(dobString) {
    var today = new Date();
    var birthDate = new Date(dobString);
    var age = today.getFullYear() - birthDate.getFullYear();
    var months = today.getMonth() - birthDate.getMonth();

    if (months < 0 || (months === 0 && today.getDate() < birthDate.getDate())) { age--; }

    return age;
}

function handleMaxLength(length) {
    var commentField = $("textarea[name=UserComments]");

    if (commentField.val().length <= commentField.attr("maxlength")) {
        var remains = commentField.attr("maxlength") - commentField.val().length;
        $("span[name=charcounter]").html("(" + remains + " characters available)");
        return true;
    } else {
        return false;
    }
}

function isInt(value) {
    return value % 1 === 0;
}

function saveMobilePreference(value) {
    $.cookie('likemobile', value, { domain: 'gatsplat.com', path: '/' });
    return true;
}
