<!DOCTYPE html>
<!--#include file="../classes/includelist.asp"-->
<%
    dim useRegistration : useRegistration = CBool(Settings(SETTING_MODULE_REGISTRATION))

    dim LimitYearForNone : LimitYearForNone = 0
    dim LimitYearForAdult : LimitYearForAdult = 1
    dim LimitYearForMinor : LimitYearForMinor = 2
    
    dim eventId, customFieldCount
    customFieldCount = 0

    if len(Request.QueryString(QUERYSTRING_VAR_EVENTID)) = 0 then
        eventId = ""
    else
        eventId = DecodeId(Request.QueryString(QUERYSTRING_VAR_EVENTID))
    end if

    dim fields
    set fields = new WaiverCustomFieldCollection
    fields.LoadAll

    for m = 0 to fields.Count - 1
        if IsNull(fields(m).Notes) OR IsEmpty(fields(m).Notes) OR Len(fields(m).Notes) = 0 Then
            fields(m).Notes = ""
        end if
        if IsNull(fields(m).AdditionalInformation) OR IsEmpty(fields(m).AdditionalInformation) OR Len(fields(m).AdditionalInformation) = 0 Then
            fields(m).AdditionalInformation = ""
        end if
    next

dim myLegalese, printLegaleseText
set myLegalese = new WaiverLegalese
myLegalese.LoadCurrentBySite
printLegaleseText = myLegalese.LegaleseText

    Dim firstTimeThroughGeneralLocation
    firstTimeThroughGeneralLocation = true
    
%>

<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title><%= SiteInfo.Name %> Online Waiver</title>
        <meta http-equiv="cache-control" content="max-age=0" /> 
        <meta http-equiv="cache-control" content="no-cache" /> 
        <meta http-equiv="expires" content="0" /> 
        <meta http-equiv="expires" content="Tue, 01 Jan 1980 1:00:00 GMT" /> 
        <meta http-equiv="pragma" content="no-cache" />
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">

        <link type="text/css" rel="stylesheet" href="../content/css/jquery.ui.all.css" />
        <link type="text/css" rel="stylesheet" href="../content/css/jquery.mobile-1.1.0.css" />
        <link type="text/css" rel="Stylesheet" href="../content/css/jquery.fancybox-1.3.4.css" />
        <link type="text/css" rel="stylesheet" href="../content/css/jquery.powertip-light.css" />
        <link type="text/css" rel="stylesheet" href="../content/css/waiver.css?<%=ANTI_CACHE_STRING %>" />
        <link type="text/css" rel="Stylesheet" href="../content/css/eventregistration.css?<%=ANTI_CACHE_STRING %>" />

        <script type="text/javascript" src="../content/js/jquery-1.7.2.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.ui.core.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.ui.widget.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.ui.datepicker.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.mobile-1.1.0.min.js"></script>
        <!--[if lt IE 9]>
            <script type="text/javascript" src="../content/js/flashcanvas.js"></script>
        <![endif]-->
        <script type="text/javascript" src="../content/js/jSignature.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.fancybox-1.3.4.pack.js"></script>
        <script type="text/javascript" src="../content/js/jquery.powertip.min.js"></script>
        <script type="text/javascript" src="../content/js/json2.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.urlencode.js"></script>
        <script type="text/javascript" src="../content/js/searchPlayers.js?<%=ANTI_CACHE_STRING %>"></script>
        <script type="text/javascript" src="../content/js/waivernavigation.js?<%=ANTI_CACHE_STRING %>"></script>
        <script type="text/javascript" src="../content/js/waivervalidation.js?<%=ANTI_CACHE_STRING %>"></script>
        <script type="text/javascript" src="../content/js/string.prototypes.js?<%=ANTI_CACHE_STRING %>"></script>
        <script type="text/javascript" src="../content/js/common.js?<%=ANTI_CACHE_STRING %>"></script>
        <script type="text/javascript">
            $(document).ready(function () {
                $(".tooltip").powerTip({ mouseOnToPopup: true, placement: 'n', smartPlacement: true });
                initWizard();

                var today = new Date;
                var formattedDate = (today.getMonth() + 1) + "/" + today.getDate() + "/" + today.getFullYear();
                <% if useRegistration then %>
                loadCalendar(formattedDate);
                <% else %>
                loadOpenCalendar(formattedDate);
                <% end if%>

                resetForm();
<%
                If Request.QueryString("search") = "true" Then
                    response.write "showWaiverSearch();"
                End If
%>
            });


            window.onbeforeunload = function () { resetForm(); }

            function resetForm() {
                $("#txtAdultHash").val("");
                $("#txtMinor1Hash").val("");
                $("#txtMinor2Hash").val("");
                $("#txtMinor3Hash").val("");
                $("#txtMinor4Hash").val("");
                $("#txtMinor5Hash").val("");
                $("#txtMinor6Hash").val("");
                $("#txtSelfFirstName").val("");
                $("#txtSelfLastName").val("");
                $("#ddlSelfDOBMonth").val("").selectmenu("refresh");
                $("#ddlSelfDOBDay").val("").selectmenu("refresh");
                $("#ddlSelfDOBYear").val("").selectmenu("refresh");
                $("#txtSelfPhoneNumber").val("");
                $("#txtSelfEmailAddress").val("");
                $("#txtSelfConfirmEmail").val("");
                $("#txtSelfAddress").val("");
                $("#txtSelfCity").val("");
                $("#ddlSelfState").val("<%=SiteInfo.State %>").selectmenu("refresh");
                $("#txtSelfZipCode").val("");
                $("#txtParentFirstName").val("");
                $("#txtParentLastName").val("");
                $("#txtParentPhoneNumber").val("");
                $("#txtParentEmailAddress").val("");
                $("#txtParentConfirmEmail").val("");
                $("#txtParentAddress").val("");
                $("#txtParentCity").val("");
                $("#ddlParentState").val("<%=SiteInfo.State %>").selectmenu("refresh");
                $("#txtParentZipCode").val("");
                $("#txtMinor1FirstName").val("");
                $("#txtMinor1LastName").val("");
                $("#ddlMinor1DOBMonth").val("").selectmenu("refresh");
                $("#ddlMinor1DOBDay").val("").selectmenu("refresh");
                $("#ddlMinor1DOBYear").val("").selectmenu("refresh");
                $("#txtMinor2FirstName").val("");
                $("#txtMinor2LastName").val("");
                $("#ddlMinor2DOBMonth").val("").selectmenu("refresh");
                $("#ddlMinor2DOBDay").val("").selectmenu("refresh");
                $("#ddlMinor2DOBYear").val("").selectmenu("refresh");
                $("#txtMinor3FirstName").val("");
                $("#txtMinor3LastName").val("");
                $("#ddlMinor3DOBMonth").val("").selectmenu("refresh");
                $("#ddlMinor3DOBDay").val("").selectmenu("refresh");
                $("#ddlMinor3DOBYear").val("").selectmenu("refresh");
                $("#txtMinor4FirstName").val("");
                $("#txtMinor4LastName").val("");
                $("#ddlMinor4DOBMonth").val("").selectmenu("refresh");
                $("#ddlMinor4DOBDay").val("").selectmenu("refresh");
                $("#ddlMinor4DOBYear").val("").selectmenu("refresh");
                $("#txtMinor5FirstName").val("");
                $("#txtMinor5LastName").val("");
                $("#ddlMinor5DOBMonth").val("").selectmenu("refresh");
                $("#ddlMinor5DOBDay").val("").selectmenu("refresh");
                $("#ddlMinor5DOBYear").val("").selectmenu("refresh");
                $("#txtMinor6FirstName").val("");
                $("#txtMinor6LastName").val("");
                $("#ddlMinor6DOBMonth").val("").selectmenu("refresh");
                $("#ddlMinor6DOBDay").val("").selectmenu("refresh");
                $("#ddlMinor6DOBYear").val("").selectmenu("refresh");
                $("#ddlMonth").val("").selectmenu("refresh");
                $("#txtPlayDate").val("");
                $("#ddlGrouped").val("").selectmenu("refresh");
                $("#ddlEventGroup").val("").selectmenu("refresh");
                $("#ddlPlayTime").val("");
                $("#btnFinish").removeAttr("disabled");
                $("#txtSearchOverride").val('false');


                //reset custom fields, if any
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
                                switch (document.getElementById(values[i].name).type.toUpperCase()) {
                                    case 'SELECT':
                                        $("#" + values[i].name).val("").selectmenu("refresh");
                                        break;
                                    case 'CHECKBOX':
                                        $("#" + values[i].name).attr('checked', false); 
                                        break;
                                    default:
                                        $("#" + values[i].name).val("");
                                }
                        }
                    }
            }

            function showAdditionalInfo(customField) {
                $.fancybox({
                    'content': $(customField).data("info"),
                    'showCloseButton': true
                });
            }
        </script>
    </head>
    <body>
        <noscript style="font-size: 20px; color: red; text-align: center" >
            Javascript must be enabled to use this form.<br/>To enable Javascript, follow <a href="https://www.gatsplat.com/javascript.asp">These Instructions.</a>
            <br/><br/>
        </noscript>

        <form id="myForm" action="save.asp" method="post" data-ajax="false">
            <input type="hidden" id="txtEventId" name="txtEventId" value="<%= eventId %>" />
            <input type="hidden" id="txtSignature" name="txtSignature" value="" />
            <input type="hidden" id="txtSelectedPath" name="txtSelectedPath" value="" />
            <input type="hidden" id="txtNumberOfMinors" name="txtNumberOfMinors" value="" />
            <input type="hidden" id="txtMinAge" value="<%= Settings(SETTING_WAIVER_MINAGE) %>" />
            <input type="hidden" id="txtUseRegistration" name="txtUseRegistration" value="<%= useRegistration %>" />
            <input type="hidden" id="txtCheckingIn" name="txtCheckingIn" value="" />
            <input type="hidden" id="txtCheckInPID" name="txtCheckInPID" value="" />
            <input type="hidden" id="txtSearchOverride" name="txtSearchOverride" value="false" />

            <div id="container" style="margin: 0 auto; text-align: center; width: 740px">

<!-- Waiver Prompt/Search -->

                <div id="divPrompt"  class="waiverSection" style="text-align: center; display: none;">
                    <h2>Have you ever filled out an online waiver with us before?</h2>
                    <table style="margin: 0 auto; width: 719px">
                        <tr>
                            <td colspan="2" style="text-align: center; width: 240px">
                                <button type="button"  onClick="showWaiverSearch();">Yes</button>
                            </td>
                            <td colspan="2" style="text-align: center; width: 240px">
                                <button type="button"  onClick="startNewWaiver();">No</button>
                            </td>
                            <td colspan="2" style="text-align: center; width: 240px">
                                <button type="button"  onClick="showWaiverSearch();">Not Sure</button>
                            </td>
                        </tr>
                    </table>
                </div>

                <div id="divMessage1"  class="waiverSection" style="text-align: center; display: none;">
                    <h2>Possible Player Conflict</h2>
                    <table style="margin: 0 auto; width: 719px">
                        <tr>
                            <td style="text-align: center; width: 240px">
                                It appears that you may already have a waiver on file.<br />Please use the form below to help us determine your current waiver status.
                            </td>
                        </tr>
                    </table>
                </div>
                            
                <div id="divSearchNameDOB" class="waiverSection" style="text-align: center; display: none;">
                    <h2>Search Player Waiver</h2>
                    <table class="blockCenter">
                        <tr>
                            <td class="inlineRight" style="width: 105px">First Name</td>
                            <td>
                                <input type="text" maxlength="50" id="txtSearchFirstName" name="txtSearchFirstName" value="" onBlur="this.value = this.value.trim(); validateRequired('txtSearchFirstName','valSearchFirstNameReq')" />
                                <span style="color: red; display: none" id="valSearchFirstNameReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Last Name:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtSearchLastName" name="txtSearchLastName" value="" onBlur="this.value = this.value.trim(); validateRequired('txtSearchLastName','valSearchLastNameReq')" />
                                <span style="color: red; display: none" id="valSearchLastNameReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Date of Birth:</td>
                            <td>
                                <fieldset data-role="controlgroup" data-type="horizontal">
                                    <% BuildYearDropDownList "ddlSearchDOBDay", "ddlSearchDOBMonth", "ddlSearchDOBYear", LimitYearForNone %>
                                    <% BuildMonthDropDownList "ddlSearchDOBDay", "ddlSearchDOBMonth", "ddlSearchDOBYear" %>
                                    <% BuildDayDropDownList "ddlSearchDOBDay", "ddlSearchDOBMonth", "ddlSearchDOBYear" %>
                                </fieldset>
                                
                                <span style="color: red; display: none" id="valSearchDOBReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <span id="valSearchResults" style="color: red; display: none"></span>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <button type="button" id="btnFindWaiver" onClick="return findWaiver(0)" onDblClick="return false;">Search</button>
                                <button type="button" id="btnCancelSearch" onClick="initWizard();" onDblClick="return false;">Cancel</button>
                            </td>
                        </tr>
                    </table>
                </div>

                <div id="divSearchResults" class="waiverSection" style="text-align: center; display: none;">
                    <h2>Search Results</h2>
                    <table class="blockCenter">
                        <tr>
                            <td style="text-align: center;">
                                 <table style="text-align: center;">
                                 <tr>
                                     <td style="font-weight:bold">We found your waiver!<br /></td>
                                 </tr>
                                 <tr>
                                     <td>Would you like to Check-In now?</td>
                                 </tr>
                                 <tr>
                                     <td><button type="button" id="btnCheckIn" onClick="checkIn();" onDblClick="return false;">Yes</button>
                                 </tr>
                                 </table>
                            </td>
                        </tr>
                    </table>
                </div>
                <div id="divNoSearchResults" class="waiverSection" style="text-align: center; display: none;">
                    <h2>Search Results</h2>
                    <table class="blockCenter">
                        <tr>
                            <td style="text-align: center;">
                                <table style="text-align: center;">
                                <tr>
                                    <td style="font-weight:bold; text-align: center;">Sorry, we couldn't find your current waiver.<br /></td>
                                </tr>
                                <tr>
                                    <td style="text-align: left;">Please check the player's information above and try again. If you are still unable to find your waiver, please try one of the following:<br /><br /></td>
                                </tr>
                                <tr>
                                    <td  style="text-align: center;">
                                        <table style="width: 600px; text-align: center;">
                                        <tr>
                                            <td style="width: 100px">&nbsp;</td>
                                            <td style="text-align: left;">If you know the email you used, we can send you an email for any reservations you have submitted.</td>
                                            <td style="text-align: center;"><button type="button" id="btnEmailMe" onClick="emailMeSearchWaiver();" onDblClick="return false;">Email Me</button></td>
                                        </tr>
                                        <tr>
                                            <td colspan="3" style="font-weight:bold;text-align: center;"><br /> OR </td>
                                        </tr>
                                        <tr>
                                             <td style="width: 100px">&nbsp;</td>
                                             <td style="text-align: left;">You can fill out a new waiver.</td>
                                            <td style="text-align: center;"><button type="button" id="btnNewWaiver" onClick="startNewWaiverFromSearch();" onDblClick="return false;">New Waiver</button></td>
                                        </tr>
                                        </table>
                                    </td>
                                </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </div>

                <div id="divEmailMeSearchWaiver" class="waiverSection" style="text-align: center; display: none;">
                    <h2>Email My Reservation</h2>
                    <table class="blockCenter">
                        <tr>
                            <td style="text-align: center;">
                                <table style="text-align: center;">
                                <tr>
                                    <td colspan="2">Please enter the email address you provided when you filled out your last waiver.<br /></td>
                                </tr>
                                <tr>
                                    <td class="inlineRight">Email:</td>
                                    <td>
                                        <input type="text" maxlength="200" id="txtWaiverSearchEmailAddress" name="txtWaiverSearchEmailAddress" onBlur="this.value = this.value.trim(); validateEmail('txtWaiverSearchEmailAddress','valWaiverSearchEmailAddressReq','valWaiverSearchEmailAddressFormat')" />
                                    </td>
                                </tr>
                                <tr>
                                    <td />
                                    <td>
                                        <span style="color: red; display: none" id="valWaiverSearchEmailAddressReq">Required</span>
                                        <span style="color: red; display: none" id="valWaiverSearchEmailAddressFormat">Invalid email address</span>
                                    </td>
                                </tr>
                                <tr>
                                    <td colspan="2" style="text-align: center;"><button type="button" id="btnWaiverSearchEmailAddress" onClick="sendWaiverSearchEmail('txtWaiverSearchEmailAddress');" onDblClick="return false;">Send</button></td>
                                </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </div>
                <div id="divEmailMeSearchWaiverResults" class="waiverSection" style="text-align: center; display: none;">
                    <h2>Email My Reservation</h2>
                    <table class="blockCenter">
                        <tr>
                            <td style="text-align: center;">
                                <div id="divEmailMeSearchWaiverResultsDisplay">&nbsp;</div>
                            </td>
                        </tr>
                    </table>
                </div>
                <div id="divEmailMeSearchWaiverNOResults" class="waiverSection" style="text-align: center; display: none;">
                    <h2>Email My Reservation</h2>
                    <table class="blockCenter">
                        <tr>
                            <td style="text-align: center;">
                                <table style="text-align: center;">
                                <tr>
                                    <td style="text-align: center;">Sorry,<br /> we couldn't find any current waivers associated with that email address.<br /></td>
                                </tr>
                                <tr>
                                    <td  style="text-align: center;">
                                        <table style="width: 600px; text-align: center;">
                                        <tr>
                                            <td style="width: 100px">&nbsp;</td>
                                            <td colspan="2" style="text-align: left;">Here are some options you can try:<br /><br /></td>
                                        </tr>                                        
                                        <tr>
                                            <td style="width: 100px">&nbsp;</td>
                                            <td style="text-align: left;">Try searching using a different email addess.</td>
                                            <td style="text-align: center;"><button type="button" id="Button1" onClick="emailMeSearchWaiver();" onDblClick="return false;">Email Me</button></td>
                                        </tr>
                                        <tr>
                                            <td colspan="3" style="font-weight:bold;text-align: center;"><br /> OR </td>
                                        </tr>
                                        <tr>
                                             <td style="width: 100px">&nbsp;</td>
                                             <td style="text-align: left;">You can fill out a new waiver.</td>
                                            <td style="text-align: center;"><button type="button" id="Button2" onClick="startNewWaiverFromSearch();" onDblClick="return false;">New Waiver</button></td>
                                        </tr>
                                        <tr>
                                            <td colspan="3" style="font-weight:bold;text-align: center;"><br /> OR </td>
                                        </tr>
                                        <tr>
                                             <td style="width: 100px">&nbsp;</td>
                                             <td style="text-align: left;">You can contact us during our normal business hours and a representative can assist you.</td>
                                            <td style="text-align: center;"></td>
                                        </tr>
                                        </table>
                                    </td>
                                </tr>
                                </table>
                            </td>
                        </tr>
                    </table>
                </div>
                
                <div id="divStreetChooser" class="waiverSection" style="text-align: center; display: none;">
                    <h2>Street Chooser</h2>
                    <table class="blockCenter">
                                <tr>
                                    <td>Please select the street of the address you provided when you filled out the waiver.</td>
                                </tr>
                                <tr>
                                    <td style="text-align: center;">
                                        <div id="divStreetsRadios" style="width: 250px; margin: 0 auto">&nbsp;</div>    
                                    </td>
                                </tr> 
                                <tr>
                                    <td style="text-align: center;"><button type="button" id="Button3" onClick="streetChosen(0);" onDblClick="return false;">OK</button></td>
                                </tr>
                                </table>
                </div>

                <div id="divInvalidWaiver" class="waiverSection" style="text-align: center; display: none;">
                    <h2>Expired waiver</h2>
                    <table class="blockCenter">
                                <tr>
                                    <td>Sorry, but your waiver has expired.  Please click the button below to sign an updated waiver.</td>
                                </tr>
                                <tr>
                                    <td style="text-align: center;">
                                        <div id="divNewExpiredWaiver" style="width: 250px; margin: 0 auto"><button name="btnNewExpiredWaiver" type="button" value="" onClick="startNewWaiverFromSearch();">OK</button> </div>    
                                    </td>
                                </tr>
                                </table>
                </div>
<!-- End Waiver Prompt/Search -->
        <div id="divWaiver" style="text-align: center; display: block;">
                <div id="divLegalese" style="text-align: center; display: none;">
                    <div style="width: 650px; margin: 15px auto; border: 1px solid grey; text-align:left;">
                        <blockquote> 
                            <% if month(now()) >= 11 then %>
                            <h3 style="color: red; text-align: center">Remember, waivers are good for the calendar year.  If playing after Jan 1, you will need to do a new waiver then.</h3>
                            <% end if %>

                            <center>
                                <a href="<%= SiteInfo.HomeUrl %>">
                                    <img src="<%= LOGO_URL %>" border="0" alt="Go to <%= SiteInfo.Name %> Home Page."  title="Go to <%= SiteInfo.Name %> Home Page.">
                                </a>
                            </center>  

                            <%= printLegaleseText %>
                            </blockquote>
                    </div>

                    <blockquote style="text-align:left">By hitting accept, you are consenting to the use of your electronic signature in lieu of an original signature on paper. You have the right to request that you sign a paper copy instead which is available at our location. By hitting accept, you are waiving that right. After consent, you may, upon written request to us, obtain a paper copy of an electronic record. No fee will be charged for such copy and no special hardware or software is required to view it. Your agreement to use an electronic signature with us for any documents will continue until such time as you notify us in writing that you no longer wish to use an electronic signature. There is no penalty for withdrawing your consent. You should always make sure that we have a current email address in order to contact you regarding any changes, if necessary.</blockquote>

                    <table style="margin: auto" >
                        <tr>
                            <td>
                                <button type="button" onClick="navigateNext()">Accept</button>
                            </td>
                            <td>
                                <button type="button"  onClick="window.location.href = '<%= SiteInfo.HomeUrl %>'">Decline</button>
                            </td>
                        </tr>
                    </table>
                </div>


                <div id="divParicipantSelection" class="waiverSection" style="text-align: center; display: none;">
                    <img src="<%= LOGO_URL %>" /> <h2>Who are you filling out waivers for?</h2>
                    <table style="margin: 0 auto; width: 719px">
                        <tr>
                            <td colspan="2" style="text-align: center; width: 240px">
                                <button type="button"  onClick="setupWizardPath(0);">Adult</button>
                            </td>
                            <td colspan="2" style="text-align: center; width: 240px;">
                                <button type="button"  onClick="setupWizardPath(1);">Minor(s)</button>
                            </td>
                            <td colspan="2" style="text-align: center; width: 240px">
                                <button type="button"  onClick="setupWizardPath(2);">Adult &amp; Minor(s)</button>
                            </td>
                        </tr>
                        <tr id="rowMinors" style="display: none">
                            <td>
                                <button type="button"  onClick="setNumberOfMinors(1)">1 Minor</button>
                            </td>
                            <td>
                                <button type="button"  onClick="setNumberOfMinors(2)">2 Minors</button>
                            </td>
                            <td>
                                <button type="button"  onClick="setNumberOfMinors(3)">3 Minors</button>
                            </td>
                            <td>
                                <button type="button"  onClick="setNumberOfMinors(4)">4 Minors</button>
                            </td>
                            <td>
                                <button type="button"  onClick="setNumberOfMinors(5)">5 Minors</button>
                            </td>
                            <td>
                                <button type="button"  onClick="setNumberOfMinors(6)">6 Minors</button>
                            </td>
                        </tr>
                    </table>
                </div>


                <div id="divSelf" class="waiverSection" style="text-align: center; display: none;">
                    <h2>Adult Information</h2>
                    <table class="blockCenter">
                        <tr>
                            <td class="inlineRight" style="width: 100px">First Name</td>
                            <td>
                                <input type="text" maxlength="50" id="txtSelfFirstName" name="txtSelfFirstName" value="" onBlur="this.value = this.value.trim(); validateRequired('txtSelfFirstName','valSelfFirstNameReq')" />
                                <span style="color: red; display: none" id="valSelfFirstNameReq">Required</span>
                            </td>
                            <td />
                        </tr>
                        <tr>
                            <td class="inlineRight">Last Name:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtSelfLastName" name="txtSelfLastName" onBlur="this.value = this.value.trim(); validateRequired('txtSelfLastName','valSelfLastNameReq')" />
                                <span style="color: red; display: none" id="valSelfLastNameReq">Required</span>
                            </td>
                            <td />
                        </tr>
                        <tr>
                            <td class="inlineRight">Date of Birth:</td>
                            <td>
                                <fieldset data-role="controlgroup" data-type="horizontal">
                                    <% BuildYearDropDownList "ddlSelfDOBDay", "ddlSelfDOBMonth", "ddlSelfDOBYear", LimitYearForAdult %>
                                    <% BuildMonthDropDownList "ddlSelfDOBDay", "ddlSelfDOBMonth", "ddlSelfDOBYear" %>
                                    <% BuildDayDropDownList "ddlSelfDOBDay", "ddlSelfDOBMonth", "ddlSelfDOBYear" %>
                                </fieldset>
                            
                                <span style="color: red; display: none" id="valSelfDOBReq">Required</span>
                                <span style="color: red; display: none" id="valSelfDBOVerification">Must be at least 18</span>
                            </td>
                            <td />
                        </tr>
                        <tr>
                            <td class="inlineRight">Phone:</td>
                            <td>
                                <input type="text" maxlength="20" id="txtSelfPhoneNumber" name="txtSelfPhoneNumber" onBlur="this.value = this.value.trim(); validatePhoneNumber('txtSelfPhoneNumber','valSelfPhoneNumberReq','valSelfPhoneNumberFormat')" />
                            </td>
                            <td />
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valSelfPhoneNumberReq">Required</span>
                                <span style="color: red; display: none" id="valSelfPhoneNumberFormat">Please use (999)999-9999 format</span>
                            </td>
                            <td />
                        </tr>
                        <tr>
                            <td class="inlineRight">Email:</td>
                            <td>
                                <input type="text" maxlength="200" id="txtSelfEmailAddress" name="txtSelfEmailAddress" onBlur="this.value = this.value.trim(); validateEmail('txtSelfEmailAddress','valSelfEmailAddressReq','valSelfEmailAddressFormat')" />
                            </td>
                            <td />
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valSelfEmailAddressReq">Required</span>
                                <span style="color: red; display: none" id="valSelfEmailAddressFormat">Invalid email address</span>
                            </td>
                            <td />
                        </tr>
                        <tr style="display: none">
                            <td class="inlineRight">Confirm Email:</td>
                            <td>
                                <input type="text" maxlength="200" id="txtSelfConfirmEmail" name="txtSelfConfirmEmail" onBlur="this.value = this.value.trim(); validateConfirmEmail('txtSelfEmailAddress','txtSelfConfirmEmail','valSelfConfirmEmailMatch')" />
                            </td>
                            <td />
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valSelfConfirmEmailMatch">Email address does not match</span>
                            </td>
                            <td />
                        </tr>
                        <tr>
                            <td class="inlineRight">Address:</td>
                            <td>
                                <input type="text" maxlength="100" id="txtSelfAddress" name="txtSelfAddress" onBlur="this.value = this.value.trim(); validateRequired('txtSelfAddress','valSelfAddressReq')"  />
                            </td>
                            <td />
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valSelfAddressReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valStreetReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">City:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtSelfCity" name="txtSelfCity" value="" onBlur="this.value = this.value.trim(); validateRequired('txtSelfCity','valSelfCityReq')" />
                            </td>
                            <td />
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valSelfCityReq">Required</span>
                            </td>
                            <td />
                        </tr>
                        <tr>
                            <td class="inlineRight"><% 
                                        select case lcase(SiteInfo.Country)
                                            case "ca"
                                                response.Write "Province:"
                                            case "us"
                                                response.Write "State:"
                                        end select
                                    %></td>
                            <td>
                                <% BuildStateDropDownList "ddlSelfState","", SiteInfo.Country %>
                            </td>
                            <td />
                        </tr>
                        <tr>
                            <td class="inlineRight"><% 
                                        select case lcase(SiteInfo.Country)
                                            case "ca"
                                                response.Write "Postal Code:"
                                            case "us"
                                                response.Write "Zip Code:"
                                        end select
                                    %></td>
                            <td>
                                <input type="text" maxlength="10" id="txtSelfZipCode" name="txtSelfZipCode" onBlur="this.value = this.value.trim(); validateRequired('txtSelfZipCode','valSelfZipCodeReq')" />
                            </td>
                            <td />
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valSelfZipCodeReq">Required</span>
                            </td>
                            <td />
                        </tr>
                        <% BuildWaiverCustomFields "Self" %>
                    </table>
                </div>


                <div id="divGuardian" class="waiverSection" style="text-align: center; display: none;">
                    <h2>Parent/Guardian Information</h2>
                    <table class="blockCenter">
                        <tr>
                            <td class="inlineRight" style="width: 110px">First Name:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtParentFirstName" name="txtParentFirstName" onBlur="this.value = this.value.trim(); validateRequired('txtParentFirstName','valParentFirstNameReq')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valParentFirstNameReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Last Name:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtParentLastName" name="txtParentLastName" onBlur="this.value = this.value.trim(); validateRequired('txtParentLastName','valParentLastNameReq')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valParentLastNameReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Phone:</td>
                            <td>
                                <input type="text" maxlength="20" id="txtParentPhoneNumber" name="txtParentPhoneNumber" value="" onBlur="this.value = this.value.trim(); validatePhoneNumber('txtParentPhoneNumber','valParentPhoneNumberReq','valParentPhoneNumberFormat')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valParentPhoneNumberReq">Required</span>
                                <span style="color: red; display: none" id="valParentPhoneNumberFormat">Please use (999)999-9999 format</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Email:</td>
                            <td>
                                <input type="text" maxlength="200" id="txtParentEmailAddress" name="txtParentEmailAddress" onBlur="this.value = this.value.trim(); validateEmail('txtParentEmailAddress','valParentEmailAddressReq','valParentEmailAddressFormat')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valParentEmailAddressReq">Required</span>
                                <span style="color: red; display: none" id="valParentEmailAddressFormat">Invalid email address</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Confirm Email:</td>
                            <td>
                                <input type="text" maxlength="200" id="txtParentConfirmEmail" name="txtParentConfirmEmail" onBlur="this.value = this.value.trim(); validateConfirmEmail('txtParentEmailAddress','txtParentConfirmEmail','valParentConfirmEmail')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valParentConfirmEmail">Email address does not match</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Address:</td>
                            <td>
                                <input type="text" maxlength="100" id="txtParentAddress" name="txtParentAddress" onBlur="this.value = this.value.trim(); validateRequired('txtParentAddress','valParentAddressReq')"  />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valParentAddressReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">City:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtParentCity" name="txtParentCity" onBlur="this.value = this.value.trim(); validateRequired('txtParentCity','valParentCityReq')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valParentCityReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">State:</td>
                            <td>
                                <% BuildStateDropDownList "ddlParentState", "", SiteInfo.Country %>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Zip:</td>
                            <td>
                                <input type="text" maxlength="10" id="txtParentZipCode" name="txtParentZipCode" onBlur="this.value = this.value.trim(); validateRequired('txtParentZipCode','valParentZipCodeReq')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valParentZipCodeReq">Required</span>
                            </td>
                        </tr>
                    </table>
                </div>


                <div id="divMinor1" class="waiverSection" style="text-align: center; display: none;">
                    <h2>First Minor's Information</h2>
                    <table class="blockCenter">
                        <tr>
                            <td class="inlineRight" style="width: 110px">First Name:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtMinor1FirstName" name="txtMinor1FirstName" onBlur="this.value = this.value.trim(); validateRequired('txtMinor1FirstName','valMinor1FirstNameReq')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor1FirstNameReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Last Name:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtMinor1LastName" name="txtMinor1LastName" onBlur="this.value = this.value.trim(); validateRequired('txtMinor1LastName','valMinor1LastNameReq')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor1LastNameReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Date of Birth:</td>
                            <td>
                                <fieldset data-role="controlgroup" data-type="horizontal">
                                    <% BuildYearDropDownList "ddlMinor1DOBDay", "ddlMinor1DOBMonth", "ddlMinor1DOBYear", LimitYearForMinor %>
                                    <% BuildMonthDropDownList "ddlMinor1DOBDay", "ddlMinor1DOBMonth", "ddlMinor1DOBYear" %>
                                    <% BuildDayDropDownList "ddlMinor1DOBDay", "ddlMinor1DOBMonth", "ddlMinor1DOBYear" %>
                                </fieldset>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor1DOBReq">Required</span>
                                <span style="color: red; display: none" id="valMinor1DOBMinimum">Must be at least <%= Settings(SETTING_WAIVER_MINAGE) %></span>
                                <span style="color: red; display: none" id="valMinor1DOBVerification">Must be under 18</span>
                            </td>
                        </tr>
                        <% BuildWaiverCustomFields "Minor1" %>
                    </table>
                </div>


                <div id="divMinor2" class="waiverSection" style="text-align: center; display: none;">
                    <h2>Second Minor's Information</h2>
                    <table class="blockCenter">
                        <tr>
                            <td class="inlineRight" style="width:110px">First Name:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtMinor2FirstName" name="txtMinor2FirstName" onBlur="this.value = this.value.trim(); validateRequired('txtMinor2FirstName','valMinor2FirstNameReq')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor2FirstNameReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Last Name:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtMinor2LastName" name="txtMinor2LastName" onBlur="this.value = this.value.trim(); validateRequired('txtMinor2LastName','valMinor2LastNameReq')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor2LastNameReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Date of Birth:</td>
                            <td>
                                <fieldset data-role="controlgroup" data-type="horizontal">
                                    <% BuildYearDropDownList "ddlMinor2DOBDay", "ddlMinor2DOBMonth", "ddlMinor2DOBYear", LimitYearForMinor %>
                                    <% BuildMonthDropDownList "ddlMinor2DOBDay", "ddlMinor2DOBMonth", "ddlMinor2DOBYear" %>
                                    <% BuildDayDropDownList "ddlMinor2DOBDay", "ddlMinor2DOBMonth", "ddlMinor2DOBYear" %>
                                </fieldset>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor2DOBReq">Required</span>
                                <span style="color: red; display: none" id="valMinor2DOBMinimum">Must be at least <%= Settings(SETTING_WAIVER_MINAGE) %></span>
                                <span style="color: red; display: none" id="valMinor2DOBVerification">Must be under 18</span>
                            </td>
                        </tr>
                        <% BuildWaiverCustomFields "Minor2" %>
                    </table>
                </div>


                <div id="divMinor3" class="waiverSection" style="text-align: center; display: none;">
                    <h2>Third Minor's Information</h2>
                    <table class="blockCenter">
                        <tr>
                            <td class="inlineRight" style="width: 110px">First Name:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtMinor3FirstName" name="txtMinor3FirstName" onBlur="this.value = this.value.trim(); validateRequired('txtMinor3FirstName', 'valMinor3FirstNameReq')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor3FirstNameReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Last Name:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtMinor3LastName" name="txtMinor3LastName" onBlur="this.value = this.value.trim(); validateRequired('txtMinor3LastName','valMinor3LastNameReq')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor3LastNameReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Date of Birth:</td>
                            <td>
                                <fieldset data-role="controlgroup" data-type="horizontal">
                                    <% BuildYearDropDownList "ddlMinor3DOBDay", "ddlMinor3DOBMonth", "ddlMinor3DOBYear", LimitYearForMinor %>
                                    <% BuildMonthDropDownList "ddlMinor3DOBDay", "ddlMinor3DOBMonth", "ddlMinor3DOBYear" %>
                                    <% BuildDayDropDownList "ddlMinor3DOBDay", "ddlMinor3DOBMonth", "ddlMinor3DOBYear" %>
                                </fieldset>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor3DOBReq">Required</span>
                                <span style="color: red; display: none" id="valMinor3DOBMinimum">Must be at least <%= Settings(SETTING_WAIVER_MINAGE) %></span>
                                <span style="color: red; display: none" id="valMinor3DOBVerification">Must be under 18</span>
                            </td>
                        </tr>
                        <% BuildWaiverCustomFields "Minor3" %>
                    </table>
                </div>


                <div id="divMinor4" class="waiverSection" style="text-align: center; display: none;">
                    <h2>Fourth Minor's Information</h2>
                    <table class="blockCenter">
                        <tr>
                            <td class="inlineRight" style="width: 110px">First Name:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtMinor4FirstName" name="txtMinor4FirstName" onBlur="this.value = this.value.trim(); validateRequired('txtMinor4FirstName','valMinor4FirstNameReq')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor4FirstNameReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Last Name:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtMinor4LastName" name="txtMinor4LastName" onBlur="this.value = this.value.trim(); validateRequired('txtMinor4LastName','valMinor4LastNameReq')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor4LastNameReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Date of Birth:</td>
                            <td>
                                <fieldset data-role="controlgroup" data-type="horizontal">
                                    <% BuildYearDropDownList "ddlMinor4DOBDay", "ddlMinor4DOBMonth", "ddlMinor4DOBYear", LimitYearForMinor %>
                                    <% BuildMonthDropDownList "ddlMinor4DOBDay", "ddlMinor4DOBMonth", "ddlMinor4DOBYear" %>
                                    <% BuildDayDropDownList "ddlMinor4DOBDay", "ddlMinor4DOBMonth", "ddlMinor4DOBYear" %>
                                </fieldset>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor4DOBReq">Required</span>
                                <span style="color: red; display: none" id="valMinor4DOBMinimum">Must be at least <%= Settings(SETTING_WAIVER_MINAGE) %></span>
                                <span style="color: red; display: none" id="valMinor4DOBVerification">Must be under 18</span>
                            </td>
                        </tr>
                        <% BuildWaiverCustomFields "Minor4" %>
                    </table>
                </div>


                <div id="divMinor5" class="waiverSection" style="text-align: center; display: none;">
                    <h2>Fifth Minor's Information</h2>
                    <table class="blockCenter">
                        <tr>
                            <td class="inlineRight" style="width: 110px">First Name:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtMinor5FirstName" name="txtMinor5FirstName" onBlur="this.value = this.value.trim(); validateRequired('txtMinor5FirstName','valMinor5FirstNameReq')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor5FirstNameReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Last Name:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtMinor5LastName" name="txtMinor5LastName" onBlur="this.value = this.value.trim(); validateRequired('txtMinor5LastName','valMinor5LastNameReq')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor5LastNameReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Date of Birth:</td>
                            <td>
                                <fieldset data-role="controlgroup" data-type="horizontal">
                                    <% BuildYearDropDownList "ddlMinor5DOBDay", "ddlMinor5DOBMonth", "ddlMinor5DOBYear", LimitYearForMinor %>
                                    <% BuildMonthDropDownList "ddlMinor5DOBDay", "ddlMinor5DOBMonth", "ddlMinor5DOBYear"  %>
                                    <% BuildDayDropDownList "ddlMinor5DOBDay", "ddlMinor5DOBMonth", "ddlMinor5DOBYear" %>
                                </fieldset>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor5DOBReq">Required</span>
                                <span style="color: red; display: none" id="valMinor5DOBMinimum">Must be at least <%= Settings(SETTING_WAIVER_MINAGE) %></span>
                                <span style="color: red; display: none" id="valMinor5DOBVerification">Must be under 18</span>
                            </td>
                        </tr>
                        <% BuildWaiverCustomFields "Minor5" %>
                    </table>
                </div>


                <div id="divMinor6" class="waiverSection" style="text-align: center; display: none;">
                    <h2>Sixth Minor's Information</h2>
                    <table class="blockCenter">
                        <tr>
                            <td class="inlineRight" style="width: 110px">First Name:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtMinor6FirstName" name="txtMinor6FirstName" onBlur="this.value = this.value.trim(); validateRequired('txtMinor6FirstName','valMinor6FirstNameReq')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor6FirstNameReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Last Name:</td>
                            <td>
                                <input type="text" maxlength="50" id="txtMinor6LastName" name="txtMinor6LastName" onBlur="this.value = this.value.trim(); validateRequired('txtMinor6LastName', 'valMinor6LastNameReq')" />
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor6LastNameReq">Required</span>
                            </td>
                        </tr>
                        <tr>
                            <td class="inlineRight">Date of Birth:</td>
                            <td>
                                <fieldset data-role="controlgroup" data-type="horizontal">
                                    <% BuildYearDropDownList "ddlMinor6DOBDay", "ddlMinor6DOBMonth", "ddlMinor6DOBYear", LimitYearForMinor %>
                                    <% BuildMonthDropDownList "ddlMinor6DOBDay", "ddlMinor6DOBMonth", "ddlMinor6DOBYear" %>
                                    <% BuildDayDropDownList "ddlMinor6DOBDay", "ddlMinor6DOBMonth", "ddlMinor6DOBYear" %>
                                </fieldset>
                            </td>
                        </tr>
                        <tr>
                            <td />
                            <td>
                                <span style="color: red; display: none" id="valMinor6DOBReq">Required</span>
                                <span style="color: red; display: none" id="valMinor6DOBMinimum">Must be at least <%= Settings(SETTING_WAIVER_MINAGE) %></span>
                                <span style="color: red; display: none" id="valMinor6DOBVerification">Must be under 18</span>
                            </td>
                        </tr>
                        <% BuildWaiverCustomFields "Minor6" %>
                    </table>
                </div>


                <div id="divPlayDate" class="waiverSection" style="display: none">
                    <h2>When are you playing?</h2>
                    <table class="blockCenter">
                        <tr>
                            <td class="inlineRight">Selected Date:</td>
                            <td>
                                <input type="text" id="txtPlayDate" name="txtPlayDate" readonly="readonly" onClick="showCalendar()" />
                                <span id="valPlayDateRequired" style="color: red; display: none">Required</span>
                            </td>
                        </tr>
                        <tr id="rowMonth" style="display: none">
                            <td class="inlineRight">Selected Month:</td>
                            <td>
                                <% if useRegistration then %>
                                <select id="ddlMonth" onChange="loadCalendar(this.value)">
                                <% else %>
                                <select id="ddlMonth" onChange="loadOpenCalendar(this.value)">
                                <% end if %>
                                    <% BuildMonthList %>
                                </select>
                            </td>
                            <td />
                        </tr>
                        <tr id="rowCalendar" style="display: none">
                            <td colspan="3">
                                <span id="calendar"></span>
                            </td>
                        </tr>
                    </table>
                </div>

            
                <div id="divPlayTime" class="waiverSection" style="display: none">
                    <% if useRegistration then %>
                    <h2>Are you playing with a group?</h2>
                    <table class="blockCenter" style="width: 350px">
                        <tr>
                            <td colspan="2">
                                <select id="ddlGrouped" name="ddlGrouped" onChange="toggleGroupTime(this.value)">
                                    <option />
                                    <option value="yes">Yes</option>
                                    <option value="no">No</option>
                                </select>
                                <span id="valGroupedRequired" style="color: red; display: none">Required</span>
                            </td>
                        </tr>
                        <tr id="rowGroup" style="display: none">
                            <td class="inlineRight">Group:</td>
                            <td>
                                <select id="ddlEventGroup" name="ddlEventGroup" onChange="checkGroup(this.value)"></select>
                                <span id="valEventGroupRequired" style="color: red; display: none">Required</span>
                            </td>
                        </tr>
                        <tr id="rowTime" style="display: none">
                            <td class="inlineRight">Time:</td>
                            <td>
                                <select id="ddlPlayTime" name="ddlPlayTime"></select>
                                <span id="valPlayTimeRequired" style="color: red; display: none">Required</span>
                            </td>
                        </tr>
                    </table>
                    <% else %>
                    <h2>What time are you playing?</h2>
                    <table class="blockCenter" style="width: 350px">
                        <tr>
                            <td class="inlineRight">Time:</td>
                            <td>
                                <select id="ddlPlayTime" name="ddlPlayTime"></select>
                                <span id="valPlayTimeRequired" style="color: red; display: none">Required</span>
                            </td>
                        </tr>
                    </table>
                    <% end if %>
                </div>

                        <% BuildWaiverCustomFields "General" %>

                <div id="divSigContainer" class="waiverSection" style="display: none">
                    <h2>Signature</h2>
                    <font size="1">Sign with your finger on touch screen, or by using your mouse, hold down the mouse button and move mouse to write.</font>

                    <div style="float:right; margin: 5px">
                        <a href="#clear" class="clearButton" onClick="$('#divSig').jSignature('reset');">Clear - Try Again</a>
                    </div>
                    <div style="clear:both;float: none"></div>
                    <div id="divSig" style="border: 1px solid #ccc; border-radius: 2em; background-color: white; width: 728px; height: 184px; margin: 0 auto"></div>
                    <span id="valSignatureRequired" style="color: red; display: none">Required</span>

                    <div style="margin: 15px 0"><%= Settings(SETTING_BLURB_WAIVER) %></div>

                    <table style="margin: 20px auto 0 auto;">
                        <tr style="margin-bottom: 20px; display: none">
                            <td class="inlineLeft">
                                <input type="checkbox" id="chkEmailList" name="chkEmailList" checked="checked" />
                            </td>
                            <td> &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; &nbsp; Receive discounts and special offers and Free Stuff
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <span id="valFormResults" style="color: red; display: none">Please scroll up to any red highlighted fields and correct them</span>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2">
                                <button type="submit" id="btnFinish" onClick="return submitForm()" onDblClick="return false;">Finish</button>
                            </td>
                        </tr>
                    </table>
                </div>

                <div id="divFinishCheckin" style="display: none">
                    <table style="margin: 20px auto 0 auto;">
                        <tr>
                            <td colspan="2">
                                <button type="submit" id="btnCheckInFinish" onClick="return submitCheckIn()" onDblClick="return false;">Finish</button>
                                <button type="button" id="Button4" onClick="initWizard();" onDblClick="return false;">Cancel</button>
                            </td>
                        </tr>
                    </table>
                </div>

            </div>
			<br/><br/><br/>
        </form>
    </body>
</html>
<%

    sub BuildDayDropDownList(DayFieldId, MonthFieldId, YearFieldId)
        Response.Write "<select id='" & DayFieldId & "' name='" & DayFieldId &"' onchange=""validateDate('" & DayFieldId & "', '" & MonthFieldId & "', '" & YearFieldId & "')"">" & vbCrlf
        Response.Write "<option value=''>D</option>"  & vbCrlf
        
        for i = 1 to 31
            Response.Write "<option value='" & i & "'>" & i & "</option>"  & vbCrlf
        next

        Response.Write "</select>" & vbCrlf
    end sub

    sub BuildMonthDropDownList(DayFieldId, MonthFieldId, YearFieldId)
        Response.Write "<select id='" & MonthFieldId & "' name='" & MonthFieldId &"' onchange=""buildDayList('" & DayFieldId & "', this.value, $('#" & YearFieldId & "').val()); " & _
                        "validateDate('" & DayFieldId & "', '" & MonthFieldId & "', '" & YearFieldId & "')"">" & vbCrlf
        Response.Write "<option value=''>M</option>"  & vbCrlf
        
        for i = 1 to 12
            Response.Write "<option value='" & i & "'>" & i & "</option>"  & vbCrlf
        next

        Response.Write "</select>" & vbCrlf
    end sub

    sub BuildMonthList()
        dim dt
        dt = cdate(Month(Now) & "/1/" & Year(Now))

        for i = 0 to 3            
		    dim currDate
            currDate = DateAdd("m", i, dt)

            response.Write "<option value='" & FormatDateTime(currDate, vbShortDate) & "'>" & MonthName(Month(currDate)) & "</option>"
        next
    end sub
    
    sub BuildYearDropDownList(DayFieldId, MonthFieldId, YearFieldId, AdultOrMinor)
        Response.Write "<select id='" & YearFieldId & "' name='" & YearFieldId &"' onchange=""buildDayList('" & DayFieldId & "', $('#" & MonthFieldId & "').val(), this.value); " & _
                        "validateDate('" & DayFieldId & "', '" & MonthFieldId & "', '" & YearFieldId & "')"">" & vbCrlf
        Response.Write "<option value=''>Year</option>"  & vbCrlf

        dim currentYear

        select case AdultOrMinor
           case LimitYearForAdult
                currentYear = DatePart("yyyy", DateAdd("yyyy",-18,Now()))
                for i = currentYear to currentYear - 100 step -1
                    Response.Write "<option value='" & i & "'>" & i & "</option>"  & vbCrlf
                next
           case LimitYearForMinor
                currentYear = DatePart("yyyy", Now())
                for i = currentYear to currentYear - 18 step -1
                    Response.Write "<option value='" & i & "'>" & i & "</option>"  & vbCrlf
                next
           case else
                currentYear = DatePart("yyyy", Now())
                for i = currentYear to currentYear - 100 step -1
                    Response.Write "<option value='" & i & "'>" & i & "</option>"  & vbCrlf
                next
        end select    
        Response.Write "</select>" & vbCrlf
    end sub

    sub BuildWaiverCustomFields(location)
        for i = 0 to fields.Count - 1
            select case fields(i).FormLocation
                case 0 'general waiver
                    if location = "General" then
                        with response
                            if firstTimeThroughGeneralLocation then
                                firstTimeThroughGeneralLocation = false
                                .write "<div id=""divWaiverCustomField"" class=""waiverSection"" style=""text-align: center; display: none;"">"
                                .write "<h2>Additional Information</h2>"
                                .write "<table class=""blockCenter"" style=""width: 350px"">"
                            end if
                            .Write "<tr>" & vbCrLf 
                            BuildWaiverCustomInput fields(i), location
                            .Write "</tr>" & vbcrlf
                        end with
                    end if

                case 1 'adult only
                    if location = "Self" then
                        response.Write "<tr>" & vbCrLf
                        BuildWaiverCustomInput fields(i), location        
                        response.Write "</tr>" & vbCrLf
                    end if

                case 2 'minors only
                    if Left(location,5) = "Minor" then
                        response.Write "<tr>" & vbCrLf
                        BuildWaiverCustomInput fields(i), location        
                        response.Write "</tr>" & vbCrLf
                    end if

                case 3 'adults and minors
                    if Left(Location,5) = "Minor" OR location = "Self" then
                        response.Write "<tr>" & vbCrLf
                        BuildWaiverCustomInput fields(i), location        
                        response.Write "</tr>" & vbCrLf
                    end if
            end select
        next

        if NOT firstTimeThroughGeneralLocation then
            response.write "</table>" & vbCrLf
            response.write "</div>"
        end if
    end sub

    sub BuildWaiverCustomInput(singleField, location)
        customFieldCount = customFieldCount + 1

        select case singleField.CustomFieldDataType.CustomFieldDataTypeId
            case FIELDTYPE_SHORTTEXT
                BuildWaiverCustomControlTextField singleField, location
            case FIELDTYPE_LONGTEXT
                BuildWaiverCustomControlTextArea  singleField, location
            case FIELDTYPE_YESNO
                BuildWaiverCustomControlCheckBox singleField, location
            case FIELDTYPE_OPTIONS
                BuildWaiverCustomControlDropDownList singleField, location
        end select
    end sub

    sub BuildWaiverCustomControlCheckBox(CurrentField, location)
        dim fieldName : fieldName = "CustomField" & CurrentField.WaiverCustomFieldId & "_" & location & "_" & customFieldCount
        dim hasAddtionalInfo : hasAddtionalInfo = cbool(len(CurrentField.AdditionalInformation) > 0)
        dim hasNotes : hasNotes = cbool(len(CurrentField.Notes) > 0)
        dim isRequired : isRequired = CurrentField.Required
        if isRequired then fieldName = fieldName & "_req"

        dim validatorName : validatorName = "val" & fieldName

        with response
            .Write "<td class=""inlineRight"" style=""width: 105px"">" & CurrentField.Name & ":</td>" & vbCrLf
            .Write "<td>" & vbCrLf
            .Write "<input type=""checkbox"" name=""" & fieldName & """ id=""" & fieldName & """ "

            .Write "class="""
            if hasNotes then .Write " tooltip"
            if isRequired then .Write " requiredCustom"
            .Write """ "

            if hasNotes then .write """ data-powertip=""" & nullablereplace(CurrentField.Notes, """", "&quot;") & """ "
            if isRequired then .Write """ data-reqValidatorId=""" & validatorName & """"

            .Write "/></td>" & vbcrlf 
            if hasAddtionalInfo then
                .write "<td>&nbsp;&nbsp;" & vbcrlf
                .Write "<img src=""../content/images/info.gif"" class=""mouseover"" alt="""" title=""additional information"" border=""0"" data-info=""" & _
                        NullableReplace(CurrentField.AdditionalInformation, """", "&quot;") & """ onclick=""showAdditionalInfo(this)"" />" & vbCrLf
                .Write "</td>" & vbCrLf
            else
                .Write "<td/>" & vbCrLf
            end if

            .write "</tr><tr>" & vbcrlf & "<td />" & vbcrlf & "<td>" & vbCrLF
            if isRequired then .Write "<span id=""" & validatorName & """ style=""color: red; display: none"">Required</span>" & vbCrLf
            .Write "</td>" & vbCrLf & "<td/>" & vbcrlf
        end with
    end sub

    sub BuildWaiverCustomControlDropDownList(CurrentField, location)
        dim fieldName : fieldName = "CustomField" & CurrentField.WaiverCustomFieldId & "_" & location & "_" & customFieldCount
        dim hasAddtionalInfo : hasAddtionalInfo = cbool(len(CurrentField.AdditionalInformation) > 0)
        dim hasNotes : hasNotes = cbool(len(CurrentField.Notes) > 0)
        dim isRequired : isRequired = CurrentField.Required
        if isRequired then fieldName = fieldName & "_req"

        dim validatorName : validatorName = "val" & fieldName

        with response
            .Write "<td class=""inlineRight"" style=""width: 105px"">" & CurrentField.Name & ":</td>" & vbCrLf
            .Write "<td>" & vbCrLf
            .Write "<select name=""" & fieldName & """ id=""" & fieldName & """  "

            .Write "class="""
            if hasNotes then .Write " tooltip"
            if isRequired then .Write " requiredCustom"
            .Write """ "

            if hasNotes then .write """ data-powertip=""" & nullablereplace(CurrentField.Notes, """", "&quot;") & """ "
            if isRequired then .Write """ data-reqValidatorId=""" & validatorName & """"

            .Write ">" & vbCrLf

            .Write "<option></option>" & vbCrLf

            dim options
            set options = new WaiverCustomOptionCollection

            options.LoadByFieldId CurrentField.WaiverCustomFieldId

            for j = 0 to options.Count - 1
                .Write "<option value=""" & options(j).Sequence & """>" & options(j).Text & "</option>" & vbCrLf
            next

            .Write "</select></td>" & vbcrlf 
            if hasAddtionalInfo then
                .write "<td>&nbsp;&nbsp;" & vbcrlf
                .Write "<img src=""../content/images/info.gif"" class=""mouseover"" alt="""" title=""additional information"" border=""0"" data-info=""" & _
                        NullableReplace(CurrentField.AdditionalInformation, """", "&quot;") & """ onclick=""showAdditionalInfo(this)"" />" & vbCrLf
                .Write "</td>" & vbCrLf
            else
                .Write "<td/>" & vbCrLf
            end if

            .write "</tr><tr>" & vbcrlf & "<td />" & vbcrlf & "<td>" & vbCrLF
            if isRequired then .Write "<span id=""" & validatorName & """ style=""color: red; display: none"">Required</span>" & vbCrLf
            .Write "</td>" & vbCrLf & "<td/>" & vbcrlf
        end with
    end sub 

    sub BuildWaiverCustomControlTextArea(CurrentField, location)
        dim fieldName : fieldName = "CustomField" & CurrentField.WaiverCustomFieldId & "_" & location & "_" & customFieldCount
        dim hasAddtionalInfo : hasAddtionalInfo = cbool(len(CurrentField.AdditionalInformation) > 0)
        dim hasNotes : hasNotes = cbool(len(CurrentField.Notes) > 0)
        dim isRequired : isRequired = CurrentField.Required
        if isRequired then fieldName = fieldName & "_req"

        dim validatorName : validatorName = "val" & fieldName

        with response
            .Write "<td class=""inlineRight"" style=""width: 105px"">" & CurrentField.Name & ":</td>" & vbCrLf
            .Write "<td/>" & vbCrLf

            if hasAddtionalInfo then
                .write "<td>&nbsp;&nbsp;" & vbcrlf
                .Write "<img src=""../content/images/info.gif"" class=""mouseover"" alt="""" title=""additional information"" border=""0"" data-info=""" & _
                        NullableReplace(CurrentField.AdditionalInformation, """", "&quot;") & """ onclick=""showAdditionalInfo(this)"" />" & vbCrLf
                .Write "</td>" & vbCrLf
            else
                .Write "<td/>" & vbCrLf
            end if

            .Write "</tr>" & vbCrLf
            .Write "<tr>" & vbCrLf
            .Write "<td colspan=""3"">" & vbCrLf

            .Write "<textarea id=""" & FieldName & """ name=""" & fieldName & """ style=""resize: none; width: 98%"" "

            .Write "class="""
            if hasNotes then .Write " tooltip"
            if isRequired then .Write " requiredCustom"
            .Write """ "

            if hasNotes then .write """ data-powertip=""" & nullablereplace(CurrentField.Notes, """", "&quot;") & """ "
            if isRequired then .Write """ data-reqValidatorId=""" & validatorName & """"

            .Write "></textarea></td>" & vbcrlf 
            .write "</tr><tr>" & vbcrlf & "<td />" & vbcrlf & "<td>" & vbCrLF
            if isRequired then .Write "<span id=""" & validatorName & """ style=""color: red; display: none"">Required</span>" & vbCrLf
            .Write "</td>" & vbCrLf & "<td/>" & vbcrlf
        end with
    end sub

    sub BuildWaiverCustomControlTextField(CurrentField, location)
        dim fieldName : fieldName = "CustomField" & CurrentField.WaiverCustomFieldId & "_" & location & "_" & customFieldCount
        dim hasAddtionalInfo : hasAddtionalInfo = cbool(len(CurrentField.AdditionalInformation) > 0)
        dim hasNotes : hasNotes = cbool(len(CurrentField.Notes) > 0)
        dim isRequired : isRequired = CurrentField.Required
        if isRequired then fieldName = fieldName & "_req"

        dim validatorName : validatorName = "val" & fieldName

        with response
            .Write "<td class=""inlineRight"" style=""width: 105px"">" & CurrentField.Name & ":</td>" & vbCrLf
            .Write "<td>" & vbCrLf
            .Write "<input type=""text"" name=""" & fieldName & """ id=""" & fieldName & """ "

            .Write "class="""
            if hasNotes then .Write " tooltip"
            if isRequired then .Write " requiredCustom"
            .Write """ "

            if hasNotes then .write """ data-powertip=""" & nullablereplace(CurrentField.Notes, """", "&quot;") & """ "
            if isRequired then .Write """ data-reqValidatorId=""" & validatorName & """"

            .Write " value=""""/></td>" & vbcrlf 
            if hasAddtionalInfo then
                .write "<td>&nbsp;&nbsp;" & vbcrlf
                .Write "<img src=""../content/images/info.gif"" class=""mouseover"" alt="""" title=""additional information"" border=""0"" data-info=""" & _
                        NullableReplace(CurrentField.AdditionalInformation, """", "&quot;") & """ onclick=""showAdditionalInfo(this)"" />" & vbCrLf
                .Write "</td>" & vbCrLf
            else
                .Write "<td/>" & vbCrLf
            end if

            .write "</tr><tr>" & vbcrlf & "<td />" & vbcrlf & "<td>" & vbCrLF
            if isRequired then .Write "<span id=""" & validatorName & """ style=""color: red; display: none"">Required</span>" & vbCrLf
            .Write "</td>" & vbCrLf & "<td/>" & vbcrlf
        end with
    end sub

    function GetCurrentTimeRounded()
        dim myHour : myHour = Hour(now())
        dim myMinutes : myMinutes = Minute(now())
    
        if cint(myMinutes) < 15 then
            myMinutes = 0
        elseif cint(myMinutes) < 45 then
            myMinutes = 30
        elseif cint(myMinutes) >= 45 then
            myHour = myHour + 1
            myMinutes = 0
        end if

        dim timeString
        if myHour < 10 then 
            timeString = "0" & myHour
        else
            timeString = myHour
        end if
        
        timeString = timeString & ":"

        if myMinutes < 10 then
            timeString = timeString & "0" & myMinutes
        else
            timeString = timeString & myMinutes
        end if

        GetCurrentTimeRounded = timeString
    end function
%>