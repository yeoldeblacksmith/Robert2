<!--#include file="../../../cute/CuteEditor_Files/include_CuteEditor.asp"--> 
<!--#include file="constants/adovbs.inc"-->
<!--#include file="constants/constants.asp"-->
<!--#include file="constants/NavigationConstants.asp"-->
<!--#include file="constants/PayPalConstants.asp"-->
<!--#include file="constants/QueryStringConstants.asp"-->
<!--#include file="constants/settings/BlurbConstants.asp"-->
<!--#include file="constants/settings/EmailConstants.asp"-->
<!--#include file="constants/settings/EmailTemplateConstants.asp"-->
<!--#include file="constants/settings/RegistrationConstants.asp"-->
<!--#include file="constants/settings/UserConstants.asp"-->
<!--#include file="constants/settings/WaiverConstants.asp"-->
<!--#include file="constants/tables/AvailableDate.asp"-->
<!--#include file="constants/tables/CustomFieldDataType.asp"-->
<!--#include file="constants/tables/EventType.asp"-->
<!--#include file="constants/tables/Invitation.asp"-->
<!--#include file="constants/tables/OrderDetail.asp"-->
<!--#include file="constants/tables/Player.asp"-->
<!--#include file="constants/tables/PlayDateTime.asp"-->
<!--#include file="constants/tables/RegistrationCustomField.asp"-->
<!--#include file="constants/tables/RegistrationCustomOption.asp"-->
<!--#include file="constants/tables/RegistrationCustomValue.asp"-->
<!--#include file="constants/tables/Role.asp"-->
<!--#include file="constants/tables/ScheduledEvent.asp"-->
<!--#include file="constants/tables/ScheduledEventPayment.asp"-->
<!--#include file="constants/tables/Site.asp"-->
<!--#include file="constants/tables/SiteSettings.asp"-->
<!--#include file="constants/tables/User.asp"-->
<!--#include file="constants/tables/Waiver.asp"-->
<!--#include file="constants/tables/WaiverCustomField.asp"-->
<!--#include file="constants/tables/WaiverCustomOption.asp"-->
<!--#include file="constants/tables/WaiverCustomValue.asp"-->
<!--#include file="constants/tables/WaiverLegalese.asp"-->
<!--#include file="utilities.asp"-->


<%
'=====================================
'   DO NOT MERGE 
'=====================================    
%>
<!--#include file="MyTools.asp"-->
<%
'=====================================
'   END DO NOT MERGE 
'=====================================    
%>




<!--#include file="bl/CustomFields/CustomFieldDataType.asp"-->
<!--#include file="bl/CustomFields/CustomFieldDataTypeCollection.asp"-->
<!--#include file="bl/CustomFields/Registration/RegistrationCustomField.asp"-->
<!--#include file="bl/CustomFields/Registration/RegistrationCustomFieldCollection.asp"-->
<!--#include file="bl/CustomFields/Registration/RegistrationCustomOption.asp"-->
<!--#include file="bl/CustomFields/Registration/RegistrationCustomOptionCollection.asp"-->
<!--#include file="bl/CustomFields/Registration/RegistrationCustomValue.asp"-->
<!--#include file="bl/CustomFields/Registration/RegistrationCustomValueCollection.asp"-->
<!--#include file="bl/CustomFields/Waiver/WaiverCustomField.asp"-->
<!--#include file="bl/CustomFields/Waiver/WaiverCustomFieldCollection.asp"-->
<!--#include file="bl/CustomFields/Waiver/WaiverCustomOption.asp"-->
<!--#include file="bl/CustomFields/Waiver/WaiverCustomOptionCollection.asp"-->
<!--#include file="bl/CustomFields/Waiver/WaiverCustomValue.asp"-->
<!--#include file="bl/CustomFields/Waiver/WaiverCustomValueCollection.asp"-->
<!--#include file="bl/email/CancellationEmail.asp"-->
<!--#include file="bl/email/ConfirmationEmail.asp"-->
<!--#include file="bl/email/EventConflictEmail.asp"-->
<!--#include file="bl/email/ExceptionEmail.asp"-->
<!--#include file="bl/email/lookupemail.asp"-->
<!--#include file="bl/email/NoAddressEmail.asp"-->
<!--#include file="bl/email/PaymentFailedEmail.asp" -->
<!--#include file="bl/email/PaymentPendingEmail.asp" -->
<!--#include file="bl/email/PasswordResetHashEmail.asp" -->
<!--#include file="bl/email/ReminderEmail.asp"-->
<!--#include file="bl/email/waiveremail.asp"-->
<!--#include file="bl/playdatetime/PlayDateTime.asp"-->
<!--#include file="bl/playdatetime/PlayDateTimeCollection.asp"-->
<!--#include file="bl/player/Player.asp"-->
<!--#include file="bl/player/PlayerCollection.asp"-->
<!--#include file="bl/registration/AvailableDate.asp"-->
<!--#include file="bl/registration/ScheduledEvent.asp"-->
<!--#include file="bl/registration/ScheduledEventCollection.asp"-->
<!--#include file="bl/registration/payment/OrderDetail.asp" -->
<!--#include file="bl/registration/payment/OrderDetailCollection.asp" -->
<!--#include file="bl/registration/payment/PaymentType.asp" -->
<!--#include file="bl/registration/payment/PaymentTypeCollection.asp" -->
<!--#include file="bl/registration/payment/ScheduledEventPayment.asp"-->
<!--#include file="bl/registration/payment/ScheduledEventPaymentCollection.asp"-->
<!--#include file="bl/Reporting/GlobalStatistics.asp"-->
<!--#include file="bl/site/Site.asp"-->
<!--#include file="bl/site/SiteSetting.asp"-->
<!--#include file="bl/site/SiteSettingCollection.asp"-->
<!--#include file="bl/user/AuthorizationRole.asp"-->
<!--#include file="bl/user/AuthorizationRolecollection.asp"-->
<!--#include file="bl/user/User.asp"-->
<!--#include file="bl/user/UserCollection.asp"-->
<!--#include file="bl/utilities/ArrayList.asp"-->
<!--#include file="bl/utilities/JSON.asp"-->
<!--#include file="bl/utilities/md5.asp"-->
<!--#include file="bl/waiver/waiver.asp"-->
<!--#include file="bl/waiver/waivercollection.asp"-->
<!--#include file="bl/waiver/waiverforexport.asp"-->
<!--#include file="bl/waiverlegalese/WaiverLegalese.asp"-->


<!--#include file="dl/DataConnection.asp"-->
<!--#include file="dl/EmailConnection.asp"-->
<!--#include file="dl/CustomFields/CustomFieldDataTypeDataConnection.asp"-->
<!--#include file="dl/CustomFields/Registration/RegistrationCustomFieldDataConnection.asp"-->
<!--#include file="dl/CustomFields/Registration/RegistrationCustomOptionDataConnection.asp"-->
<!--#include file="dl/CustomFields/Registration/RegistrationCustomValueDataConnection.asp"-->
<!--#include file="dl/CustomFields/waiver/WaiverCustomFieldDataConnection.asp"-->
<!--#include file="dl/CustomFields/waiver/WaiverCustomOptionDataConnection.asp"-->
<!--#include file="dl/CustomFields/waiver/WaiverCustomValueDataConnection.asp"-->
<!--#include file="dl/playdatetime/playdatetimedataconnection.asp"-->
<!--#include file="dl/player/playerdataconnection.asp"-->
<!--#include file="dl/registration/AvailableDateDataConnection.asp"-->
<!--#include file="dl/registration/ScheduledEventDataConnection.asp"-->
<!--#include file="dl/registration/payment/OrderDetailDataConnection.asp"-->
<!--#include file="dl/registration/payment/ScheduledEventPaymentDataConnection.asp"-->
<!--#include file="dl/Reporting/GlobalStatisticsDataConnection.asp"-->
<!--#include file="dl/site/SiteDataConnection.asp"-->
<!--#include file="dl/site/SiteSettingsDataConnection.asp"-->
<!--#include file="dl/user/RoleDataConnection.asp"-->
<!--#include file="dl/user/UserDataConnection.asp"-->
<!--#include file="dl/waiver/waiverdataconnection.asp"-->
<!--#include file="dl/waiverlegalese/WaiverLegaleseDataConnection.asp"-->

<!--#include file="pl/navigationmenu.asp"-->
<!--#include file="pl/navigationmenulink.asp"-->
<!--#include file="pl/navigationmenulinkcollection.asp"-->
<!--#include file="pl/SecurityHelper.asp"-->


<%
    If Request.ServerVariables("SERVER_PORT")=80 Then ForceSSL

    dim Settings, SiteInfo, MY_SITE_GUID, ANTI_CACHE_STRING, LOGO_URL
    set SiteInfo = GetSite()

    MY_SITE_GUID = SiteInfo.SiteGuid
    ANTI_CACHE_STRING = GetAntiCacheString()
    LOGO_URL = GetLogoUrl()

    set Settings = GetSettingsDictionary()

    if cbool(Settings(SETTING_MODULE_REGISTRATION)) then SendReminders
    'if UsesPayments() then SendPaymentReminders

    ClearTempEvents

    '*********************************************************************************************
    ' clean up temporary events for the site
    '*********************************************************************************************
    private sub ClearTempEvents()
        dim temp : set temp = new ScheduledEventCollection
        temp.DeleteTemporaryEvents
    end sub

    '*********************************************************************************************
    ' redirect users on all pages to use ssl
    '*********************************************************************************************
    private sub ForceSSL()
        Dim strSecureURL : strSecureURL = "https://"
        strSecureURL = strSecureURL & Request.ServerVariables("SERVER_NAME")
        strSecureURL = strSecureURL & Request.ServerVariables("URL")
        if len(Request.QueryString) > 0 then  strSecureURL = strSecureURL & "?" & Request.QueryString
 
        Response.Redirect strSecureURL
    end sub

    private function GetAntiCacheString()
        dim text

        text = Year(Now())
        text = text & iif(Month(Now()) < 10, "0" & Month(Now()), Month(Now()))
        text = text & iif(Day(Now()) < 10, "0" & Day(Now()), Day(Now()))
        text = text & iif(Hour(Now()) < 10, "0" & Hour(Now()), Hour(Now()))
        text = text & iif(Minute(Now()) < 10, "0" & Minute(Now()), Minute(Now()))
        text = text & iif(Second(Now()) < 10, "0" & Second(Now()), Second(Now()))

        GetAntiCacheString = text
    end function

    private function GetLogoUrl()
        GetLogoUrl = left(SiteInfo.VantoraUrl, InStrRev(SiteInfo.VantoraUrl, "/")) & "logos/" & SiteInfo.DirectoryName & ".gif"
    end function

    private function GetSettingsDictionary()
        dim settings
        set settings = new SiteSettingCollection
        settings.LoadAll

        set GetSettingsDictionary = settings.ToDictionary()
    end function

    private function GetSite()
        dim urlParts
        urlParts = Split(request.ServerVariables("URL"), "/")

        dim siteName
        siteName = urlParts(2)

        dim mySite
        set mySite = new Site

        mySite.LoadByName siteName

        set GetSite = mySite
    end function

    private sub SendPaymentReminders()
        ' send reminder emails every time a page loads
        dim events
        set events = new ScheduledEventCollection
        events.LoadForPaymentReminders

        dim email
        set email = new PaymentPendingEmail

        on error resume next
        
        for i = 0 to events.Count - 1
            email.SendByObject events(i)

            if err.number = 0 then
                events(i).SavePaymentReminderDate events(i).EventId, Now()
            else
                dim exEmail : set exEmail = new ExcpetionEmail
                exEmail.Send Request, err.number, err.Description, Server.GetLastError()
            end if
        next
    end sub

    private sub SendReminders()
        ' send reminder emails every time a page loads
        dim events
        set events = new ScheduledEventCollection
        events.LoadForReminders

        dim email
        set email = new ReminderEmail

        on error resume next

        for i = 0 to events.Count - 1
            email.SendByObject events(i)

            if err.number = 0 then
                events(i).SaveReminderDate events(i).EventId, Now()
            else
                dim exEmail : set exEmail = new ExcpetionEmail
                exEmail.Send Request, err.number, err.Description, Server.GetLastError()
            end if
        next
    end sub
%>