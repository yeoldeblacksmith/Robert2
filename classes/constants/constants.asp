﻿<% 
    const AUTHORIZED_ROLE_GLOBALADMIN = "1"
    const AUTHORIZED_ROLE_ADMIN = "2"
    const AUTHORIZED_ROLE_MANAGER = "3"
    const AUTHROIZED_ROLE_USER = ""

    const COOKIE_ROLEID = "role"
    const COOKIE_SITEID = "site"
    const COOKIE_USER = "user"

    const ENCODE_SEED = 234

    const FIELDTYPE_SHORTTEXT = 1
    const FIELDTYPE_LONGTEXT = 2
    const FIELDTYPE_YESNO = 3
    const FIELDTYPE_OPTIONS = 4

    const ITEM_NAME_RESERVATIONFEE = "Reservation Deposit"

    const MIMETYPE_JSON = "application/json"
    const MIMETYPE_OTHER = "application/octet-stream"

    const PASSWORD_SALT = "@1374"

    const PAYMENTSTATUS_UNPAID = "00"
    const PAYMENTSTATUS_CONFLICTED = "03"
    const PAYMENTSTATUS_FAILED = "07"
    const PAYMENTSTATUS_PAID = "10"
    const PAYMENTSTATUS_MARKEDPAID = "12"
    const PAYMENTSTATUS_PAIDCONFLICTED = "13"
    const PAYMENTSTATUS_REFUNDED = "15"
    const PAYMENTSTATUS_CHANGEDPAYABLE = "20"
    const PAYMENTSTATUS_CHANGEDRECEIVABLE = "25"
    const PAYMENTSTATUS_CANCELLED_USER = "90"
    const PAYMENTSTATUS_CANCELLED_ADMIN = "91"
    const PAYMENTSTATUS_CANCELLED_TIMEOUT = "95"
    const PAYMENTSTATUS_TEMPORARY = "99"

    const PAYMENTTYPE_NONE = "0"
    const PAYMENTTYPE_BYEVENT = "1"
    const PAYMENTTYPE_BYPLAYER = "2"

    const SUPPORT_EMAIL_ADDRESS = "fizch26@gmail.com;rob@metalmischief.com"

    const WAIVER_VALIDITYSTATUS_GOOD = 0
    const WAIVER_VALIDITYSTATUS_EXPIRED = 1
    const WAIVER_VALIDITYSTATUS_NEWADULT = 2

    const WAIVERCUSTOMFIELDS_LOCATION_WAIVERGENERAL = 0
    const WAIVERCUSTOMFIELDS_LOCATION_ADULTONLY = 1
    const WAIVERCUSTOMFIELDS_LOCATION_MINORSONLY = 2
    const WAIVERCUSTOMFIELDS_LOCATION_ADULTANDMINORS = 3

    Dim WAIVERCUSTOMFIELDS_LOCATION_DESCRIPTION(3) 
    WAIVERCUSTOMFIELDS_LOCATION_DESCRIPTION(WAIVERCUSTOMFIELDS_LOCATION_WAIVERGENERAL) = "General Waiver"
    WAIVERCUSTOMFIELDS_LOCATION_DESCRIPTION(WAIVERCUSTOMFIELDS_LOCATION_ADULTONLY) = "Adult Info"
    WAIVERCUSTOMFIELDS_LOCATION_DESCRIPTION(WAIVERCUSTOMFIELDS_LOCATION_MINORSONLY) = "Minors Info"
    WAIVERCUSTOMFIELDS_LOCATION_DESCRIPTION(WAIVERCUSTOMFIELDS_LOCATION_ADULTANDMINORS) = "Adult and Minors Info"

    CONST INCLUDEINACTIVEWAIVERCUSTOMFIELDS = 1
    CONST EXCLUDEINACTIVEWAIVERCUSTOMFIELDS = 0



%>