function addInvite() {
    if (validateInvite()) {
        var name = escape($("#txtGuestName").val());
        var email = escape($("#txtGuestEmail").val());
        var eventId = $("#eventId").val();
        var sUrl = 'common/inviteajax.asp?act=0&id=' + eventId + "&nm=" + name + "&em=" + email;

        $.ajax({
            url: sUrl,
            error: function () { alert("problems encountered sending request"); },
            success: function () { loadTable(); }
        });
    }
}

function deleteInvite(inviteId) {
    if (confirm("Are you sure?")) {
        $.ajax({
            url: 'common/inviteajax.asp?act=5&id=' + inviteId,
            error: function () { alert("problems encountered deleting invitation"); },
            success: function () { loadTable(); }
        });
    }
}

function editInvite(inviteId) {
    $.ajax({
        url: 'common/inviteajax.asp?act=21&id=' + $("#eventId").val() + '&nm=' + inviteId,
        error: function () { alert("problems encountered"); },
        success: function (data) {
            var container = $("#datagrid");

            container.empty();
            container.html(data);
        }
    });
}

function updateInvite(inviteId) {
    if (validateInvite()) {
        var name = escape($("#txtGuestName").val());
        var email = escape($("#txtGuestEmail").val());

        $.ajax({
            url: 'common/inviteajax.asp?act=10&id=' + inviteId + "&nm=" + name + "&em=" + email,
            error: function () { alert("problems encountered sending request"); },
            success: function () { loadTable(); }
        });
    }
}

function validateInvite() {
    var valid = true;

    // validate contact email address
    var email = $("#txtGuestEmail").val().toLowerCase();
    var emailRegExp = /^([a-zA-Z0-9_\.\-\+])+\@(([a-zA-Z0-9\-])+\.)+([a-zA-Z0-9]{2,4})+$/;

    if (email.trim().length > 0) {
        if (emailRegExp.test(email)) {
            $("span[name=ValEmail]").css("display", "none");
            $("span[name=ValEmailRequired]").css("display", "none");
        } else {
            $("span[name=ValEmail]").css("display", "inline");
            valid = false;
        } // regExp test
    } else {
        $("span[name=ValEmailRequired]").css("display", "inline");
        valid = false;
    } // length test

    // validate contact name
    var name = $("#txtGuestName").val();

    if (name.trim().length == 0) {
        $("span[name=ValName]").css("display", "inline");
        valid = false;
    } else {
        $("span[name=ValName]").css("display", "none");
    }

    return valid;
}
