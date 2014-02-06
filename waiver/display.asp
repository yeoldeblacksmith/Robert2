<!--#include file="../classes/includelist.asp"-->
<%
    dim myWaiver, myWaiverCustomFields, myCustomOptions, useSvgImage, hasSingature, hasExpired, isNewAdult
    set myWaiver = new Waiver
    set myWaiverCustomFields = new WaiverCustomValueCollection
    set myCustomOptions = new WaiverCustomOptionCollection

    useSvgImage = false
    hasSingature = false
    hasExpired = false
    isNewAdult = false

    if IsNull(request.QueryString(QUERYSTRING_VAR_WAIVERID)) = false then
        myWaiver.Load request.QueryString(QUERYSTRING_VAR_WAIVERID)
        myWaiverCustomFields.LoadByWaiverId myWaiver.WaiverId
    end if

    if myWaiver.Loaded = false then
        myWaiver.PlayDate = FormatDateTime(Now(), 2)
        myWaiver.WaiverDate = FormatDateTime(Now(), 2)
        myWaiver.EmailList = true
        myWaiver.EventId = Request.QueryString(QUERYSTRING_VAR_EVENTID)
        'myWaiver.DateOfBirth = "0/0/0000"
    else
        if len(myWaiver.Signature) > 0 then
            hasSingature = true
            useSvgImage = cbool(left(myWaiver.Signature, 5) = "<?xml")
        end if

        if myWaiver.ValidityStatus = WAIVER_VALIDITYSTATUS_EXPIRED then
            hasExpired = true
        else
            if myWaiver.ValidityStatus = WAIVER_VALIDITYSTATUS_NEWADULT then
                isNewAdult = true
            end if
        end if
    end if

    dim myLegalese, printLegaleseText
    set myLegalese = new WaiverLegalese
    myLegalese.LoadCurrentBySite 
    printLegaleseText = myLegalese.LegaleseText
%>
<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
    <head>
        <title><%= SiteInfo.Name %>  Online Waiver</title>
        <meta http-equiv="Content-Type" content="text/html; charset=utf-8">
        <meta name="viewport" content="width=device-width, maximum-scale=1.0, minimum-scale=1.0" /> 

        <link type="text/css" rel="stylesheet" href="../content/css/jquery.ui.all.css" />
        <link type="text/css" rel="stylesheet" href="../content/css/jquery.mobile-1.1.0.css" />
        <link type="text/css" rel="Stylesheet" href="../content/css/jquery.fancybox-1.3.4.css" />
        <link type="text/css" rel="stylesheet" href="../content/css/waiver.css" />
        <link type="text/css" rel="Stylesheet" href="../content/css/eventregistration.css" />

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
        <script type="text/javascript" src="../content/js/json2.min.js"></script>
        <script type="text/javascript" src="../content/js/jquery.urlencode.js"></script>
        <script type="text/javascript" src="../content/js/waivernavigation.js"></script>
        <script type="text/javascript" src="../content/js/string.prototypes.js"></script>
        <script type="text/javascript" src="../content/js/common.js"></script>

        <script type="text/javascript">
            $(document).ready(function () {
                <% if useSvgImage = false then %>
                $.ajax({
                    url: "../ajax/Waiver.asp?act=20&id=<%=myWaiver.HashId%>",
                    error: function(){ alert ("problems encountered getting player signature"); },
                    success: function(data) {
                        $(".plSig").signaturePad({ lineTop: 70, drawOnly: true, output: ".plOutput", validateFields: false, displayOnly: true }).regenerate(data);
                    }
                });
                <% end if %>
                
            });
        </script>
    </head>

    <body style="background-color: white;">
        <div style="background-color: white;">
            <div id="printLink"" style="position: fixed; top: 0px; left: 0px; border-bottom: 1px solid grey; border-right: 1px solid grey; width: 200px; border-bottom-right-radius: 1.5em; background-color: white; padding: 10px">
                <a id="A1" href="#" onclick="printWaiver();return false;">
                    <img src="../content/images/printpage.jpg" alt="print" title="Print this page" style="float: left"/>
                </a>
            </div>

            <div id="divPromptContainer" style="margin: 0 auto; text-align: center; width: 740px; display: block; background-color: white;">
                <div id="divExpired"  class="waiverSection" style="text-align: center; background-color: white; <% if NOT hasExpired then response.write "display: none;" end if %> ">
                            <h1 style="text-align:center; color: red">This waiver has expired - Need one for this year</h1>
                            <table style="margin: 0 auto;">
                                <tr>
                                    <td colspan="2" style="text-align: center; width: 240px">
                                        <button type="button"  onClick="window.location = 'default.asp';">Click here to get a new waiver</button>
                                    </td>
                                </tr>
                            </table>
                  </div>
                <div id="divNewAdult"  class="waiverSection" style="text-align: center; background-color: white; <% if NOT isNewAdult then response.write "display: none;" end if %>">
                            <h1 style="text-align:center; color: red">This waiver has expired - New Adult</h1>
                            <table style="margin: 0 auto;">
                                <tr>
                                    <td colspan="2" style="text-align: center; width: 240px">
                                        <button type="button"  onClick="window.location = 'default.asp';">Click here to get a new waiver</button>
                                    </td>
                                </tr>
                            </table>
                  </div>
           </div>
            <div id="divLegalese" style="text-align: center; display: block;background-color: white;">
                <div   class="waiverSection" style="width: 740px; margin: 15px auto; border: 1px solid grey; text-align:left;">
                    <blockquote> 
                        <center>
                            <a href="<%= SiteInfo.HomeUrl %>"><img src="<%= LOGO_URL %>" border="0" alt="Go to <%= SiteInfo.Name %> Home Page."  title="Go to <%= SiteInfo.Name %> Home Page."></a> 
                        </center>  
                        <%= printLegaleseText %>
                    </blockquote>
                           
                            
                    <input type="hidden" id="txtPlayerSignature" name="txtPlayerSignature" />

                    <table style="margin: 0 auto">
                        <tr>
                            <td>Participant Name:</td>
                            <td>
                                <%= myWaiver.FirstName %>&nbsp;<%= myWaiver.LastName %>  &nbsp; 
                            </td>
                        </tr>
                        <tr>
                            <td>Participant Date of Birth:</td>
                            <td>
                                <%= FormatDateTime(myWaiver.PlayerDOB, 2) %>
                            </td>
                        </tr>
                        <% if len(myWaiver.ParentId) > 1 then %>
                        <tr>
                            <td>Parent/Guardian Name:</td>
                            <td>
                                <%= myWaiver.ParentFirstName %> &nbsp; <%= myWaiver.ParentLastName %>
                            </td>
                        </tr>
                
                        <% end if %>
                        <tr>
                            <td>Phone:</td>
                            <td>
                                <%= myWaiver.PhoneNumber %>
                            </td>
                        </tr>
                        <tr>
                            <td>Email:</td>
                            <td>
                                <%= myWaiver.EmailAddress %>
                            </td>
                        </tr>
                
                        <tr>
                            <td>Address:</td>
                            <td>
                                <%= myWaiver.Address %> &nbsp; &nbsp; <%= myWaiver.City %> , <%= myWaiver.State %> &nbsp;<%= myWaiver.ZipCode %>
                            </td>
                        </tr>
                        <%
                            For k = 0 to myWaiverCustomFields.Count -1
                                   If myWaiverCustomFields(k).LocationInWaiver <> WAIVERCUSTOMFIELDS_LOCATION_WAIVERGENERAL Then
                                         with Response
                                            .write "<tr>" & vbcrlf
                                            .write "<td>" & myWaiverCustomFields(k).Name & "</td>" & vbcrlf

                                            If myWaiverCustomFields(k).CustomFieldDataType = FIELDTYPE_OPTIONS Then
                                                 myCustomOptions.LoadByFieldId myWaiverCustomFields(k).WaiverCustomFieldId 
                                                 for m = 0 to myCustomOptions.Count -1
                                                    if CStr(myCustomOptions(m).Sequence) = myWaiverCustomFields(k).Value then
                                                        .write "<td>" & myCustomOptions(m).Text & "</td>" & vbcrlf 
                                                    end if
                                                 next
                                                 set myCustomOptions = new WaiverCustomOptionCollection
                                            Else
                                                .write "<td>" & myWaiverCustomFields(k).Value & "</td>" & vbcrlf
                                                .write "</tr>" & vbcrlf
                                            End If
                                         end with   
                                   end if
                            Next

                            For k = 0 to myWaiverCustomFields.Count -1
                                   If myWaiverCustomFields(k).LocationInWaiver = WAIVERCUSTOMFIELDS_LOCATION_WAIVERGENERAL Then
                                         with Response
                                            .write "<tr>" & vbcrlf
                                            .write "<td>" & myWaiverCustomFields(k).Name & "</td>" & vbcrlf

                                            If myWaiverCustomFields(k).CustomFieldDataType = FIELDTYPE_OPTIONS Then
                                                 myCustomOptions.LoadByFieldId myWaiverCustomFields(k).WaiverCustomFieldId 
                                                 for m = 0 to myCustomOptions.Count -1
                                                    if CStr(myCustomOptions(m).Sequence) = myWaiverCustomFields(k).Value then
                                                        .write "<td>" & myCustomOptions(m).Text & "</td>" & vbcrlf 
                                                    end if
                                                 next
                                                 set myCustomOptions = new WaiverCustomOptionCollection
                                            Else
                                                .write "<td>" & myWaiverCustomFields(k).Value & "</td>" & vbcrlf
                                                .write "</tr>" & vbcrlf
                                            End If
                                         end with   
                                   end if
                            Next
                        %>
                        <tr>
                            <td>Date Signed:</td>
                            <td>
                                <%= FormatDateTime(myWaiver.CreateDateTime, 2) %>
                            </td>
                        </tr>
                        <% if hasSingature = false then %>
                        <tr>
                            <td colspan="2" style="color: red">
                                Signature does not exist
                            </td>
                        </tr>
                        <% elseif useSvgImage then %>
                        <tr>
                            <td colspan="2">
                                Signature:<br />
                                <img src="signaturesvg.asp?id=<%= myWaiver.HashId %>" alt="Signature" />
                            </td>
                        </tr>
                        <% else %>
                        <tr>
                            <td colspan="2" class="plSig">
                                Player Signature:
                                <div class="sig sigWrapper">
                                    <div class="typed"></div>
                                    <canvas class="pad" width="498" height="105"></canvas>
                                    <input type="hidden" id="plOutput" name="plOutput" class="output"/>
                                </div>
                            </td>
                        </tr>
                        <tr>
                            <td colspan="2" class="pgSig">
                                Parent/Gaurdian Signature:
                                <div class="sig sigWrapper">
                                    <div class="typed"></div>
                                    <canvas class="pad" width="498" height="105"></canvas>
                                    <input type="hidden" id="pgOutput" name="pgOutput" class="output"/>
                                </div>
                            </td>
                        </tr>
                        <% end if %>
                    </table>
                </div>
            </div>
        </div>
    </body>
</html>
