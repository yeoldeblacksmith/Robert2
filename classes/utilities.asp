<%  
    sub BuildPaymentTypeList(SelectedId)
        dim types : set types = new PaymentTypeCollection
        types.Load

        with Response
            .Write "<select id=""PaymentType"" name=""" & SETTING_REGISTRATION_PAYMENTTYPE & """>" & vbCrLf
            
            for i = 0 to types.Count - 1
                .Write "<option value=""" & types(i).Id & """" & iif(cstr(SelectedId) = cstr(types(i).Id), "selected=""selected""", "") & ">" & types(i).Description & "</option>" & vbCrLf
            next
            
            .Write "</select>" & vbCrLf
        end with
    end sub

    sub BuildStateDropDownList(FieldId, SelectedState, Country)
        if len(SelectedState) = 0 then
            dim mySite
            set mySite = new Site
            mySite.Load MY_SITE_GUID

            SelectedState = mySite.State
        end if

        with Response
            .Write "<select id='" & FieldId & "' name='" & FieldId &"'>" & vbCrlf
            select case lcase(Country)
                case "ca"
                    .Write "<option value=""AB""" & iif(SelectedState = "AB", "selected=""selected""", "") & ">Alberta</option>" & vbCrlf
                    .Write "<option value=""BC""" & iif(SelectedState = "BC", "selected=""selected""", "") & ">British Columbia</option>" & vbCrlf
                    .Write "<option value=""MB""" & iif(SelectedState = "MB", "selected=""selected""", "") & ">Manitoba</option>" & vbCrlf
                    .Write "<option value=""NB""" & iif(SelectedState = "NB", "selected=""selected""", "") & ">New Brunswick</option>" & vbCrlf
                    .Write "<option value=""NL""" & iif(SelectedState = "NL", "selected=""selected""", "") & ">Newfoundland</option>" & vbCrlf
                    .Write "<option value=""NS""" & iif(SelectedState = "NS", "selected=""selected""", "") & ">Nova Scotia</option>" & vbCrlf
                    .Write "<option value=""NT""" & iif(SelectedState = "NT", "selected=""selected""", "") & ">Northwest Territories</option>" & vbCrlf
                    .Write "<option value=""NU""" & iif(SelectedState = "NU", "selected=""selected""", "") & ">Nunavut</option>" & vbCrlf
                    .Write "<option value=""ON""" & iif(SelectedState = "ON", "selected=""selected""", "") & ">Ontario</option>" & vbCrlf
                    .Write "<option value=""PE""" & iif(SelectedState = "PE", "selected=""selected""", "") & ">Prince Edward Island</option>" & vbCrlf
                    .Write "<option value=""QC""" & iif(SelectedState = "QC", "selected=""selected""", "") & ">Quebec</option>" & vbCrlf
                    .Write "<option value=""SK""" & iif(SelectedState = "SK", "selected=""selected""", "") & ">Saskatchewan</option>" & vbCrlf
                    .Write "<option value=""YT""" & iif(SelectedState = "YT", "selected=""selected""", "") & ">Yukon</option>" & vbCrlf
                case "us"
                    .Write "<option value=""AL""" & iif(SelectedState = "AL", "selected=""selected""", "") & ">Alabama</option>" & vbCrlf
                    .Write "<option value=""AK""" & iif(SelectedState = "AK", "selected=""selected""", "") & ">Alaska</option>" & vbCrlf
                    .Write "<option value=""AZ""" & iif(SelectedState = "AZ", "selected=""selected""", "") & ">Arizona</option>" & vbCrlf
                    .Write "<option value=""AR""" & iif(SelectedState = "AR", "selected=""selected""", "") & ">Arkansas</option>" & vbCrlf
                    .Write "<option value=""CA""" & iif(SelectedState = "CA", "selected=""selected""", "") & ">California</option>" & vbCrlf
                    .Write "<option value=""CO""" & iif(SelectedState = "CO", "selected=""selected""", "") & ">Colorado</option>" & vbCrlf
                    .Write "<option value=""CT""" & iif(SelectedState = "CT", "selected=""selected""", "") & ">Connecticut</option>" & vbCrlf
                    .Write "<option value=""DE""" & iif(SelectedState = "DE", "selected=""selected""", "") & ">Delaware</option>" & vbCrlf
                    .Write "<option value=""DC""" & iif(SelectedState = "DC", "selected=""selected""", "") & ">Dist of Columbia</option>" & vbCrlf
                    .Write "<option value=""FL""" & iif(SelectedState = "FL", "selected=""selected""", "") & ">Florida</option>" & vbCrlf
                    .Write "<option value=""GA""" & iif(SelectedState = "GA", "selected=""selected""", "") & ">Georgia</option>" & vbCrlf
                    .Write "<option value=""HI""" & iif(SelectedState = "HI", "selected=""selected""", "") & ">Hawaii</option>" & vbCrlf
                    .Write "<option value=""ID""" & iif(SelectedState = "ID", "selected=""selected""", "") & ">Idaho</option>" & vbCrlf
                    .Write "<option value=""IL""" & iif(SelectedState = "IL", "selected=""selected""", "") & ">Illinois</option>" & vbCrlf
                    .Write "<option value=""IN""" & iif(SelectedState = "IN", "selected=""selected""", "") & ">Indiana</option>" & vbCrlf
                    .Write "<option value=""IA""" & iif(SelectedState = "IA", "selected=""selected""", "") & ">Iowa</option>" & vbCrlf
                    .Write "<option value=""KS""" & iif(SelectedState = "KS", "selected=""selected""", "") & ">Kansas</option>" & vbCrlf
                    .Write "<option value=""KY""" & iif(SelectedState = "KY", "selected=""selected""", "") & ">Kentucky</option>" & vbCrlf
                    .Write "<option value=""LA""" & iif(SelectedState = "LA", "selected=""selected""", "") & ">Louisiana</option>" & vbCrlf
                    .Write "<option value=""ME""" & iif(SelectedState = "ME", "selected=""selected""", "") & ">Maine</option>" & vbCrlf
                    .Write "<option value=""MD""" & iif(SelectedState = "MD", "selected=""selected""", "") & ">Maryland</option>" & vbCrlf
                    .Write "<option value=""MA""" & iif(SelectedState = "MA", "selected=""selected""", "") & ">Massachusetts</option>" & vbCrlf
                    .Write "<option value=""MI""" & iif(SelectedState = "MI", "selected=""selected""", "") & ">Michigan</option>" & vbCrlf
                    .Write "<option value=""MN""" & iif(SelectedState = "MN", "selected=""selected""", "") & ">Minnesota</option>" & vbCrlf
                    .Write "<option value=""MS""" & iif(SelectedState = "MS", "selected=""selected""", "") & ">Mississippi</option>" & vbCrlf
                    .Write "<option value=""MO""" & iif(SelectedState = "MO", "selected=""selected""", "") & ">Missouri</option>" & vbCrlf
                    .Write "<option value=""MT""" & iif(SelectedState = "MT", "selected=""selected""", "") & ">Montana</option>" & vbCrlf
                    .Write "<option value=""NE""" & iif(SelectedState = "NE", "selected=""selected""", "") & ">Nebraska</option>" & vbCrlf
                    .Write "<option value=""NV""" & iif(SelectedState = "NV", "selected=""selected""", "") & ">Nevada</option>" & vbCrlf
                    .Write "<option value=""NH""" & iif(SelectedState = "NH", "selected=""selected""", "") & ">New Hampshire</option>" & vbCrlf
                    .Write "<option value=""NJ""" & iif(SelectedState = "NJ", "selected=""selected""", "") & ">New Jersey</option>" & vbCrlf
                    .Write "<option value=""NM""" & iif(SelectedState = "NM", "selected=""selected""", "") & ">New Mexico</option>" & vbCrlf
                    .Write "<option value=""NY""" & iif(SelectedState = "NY", "selected=""selected""", "") & ">New York</option>" & vbCrlf
                    .Write "<option value=""NC""" & iif(SelectedState = "NC", "selected=""selected""", "") & ">North Carolina</option>" & vbCrlf
                    .Write "<option value=""ND""" & iif(SelectedState = "ND", "selected=""selected""", "") & ">North Dakota</option>" & vbCrlf
                    .Write "<option value=""OH""" & iif(SelectedState = "OH", "selected=""selected""", "") & ">Ohio</option>" & vbCrlf
                    .Write "<option value=""OK""" & iif(SelectedState = "OK", "selected=""selected""", "") & ">Oklahoma</option>" & vbCrlf
                    .Write "<option value=""OR""" & iif(SelectedState = "OR", "selected=""selected""", "") & ">Oregon</option>" & vbCrlf
                    .Write "<option value=""PA""" & iif(SelectedState = "PA", "selected=""selected""", "") & ">Pennsylvania</option>" & vbCrlf
                    .Write "<option value=""RI""" & iif(SelectedState = "RI", "selected=""selected""", "") & ">Rhode Island</option>" & vbCrlf
                    .Write "<option value=""SC""" & iif(SelectedState = "SC", "selected=""selected""", "") & ">South Carolina</option>" & vbCrlf
                    .Write "<option value=""SD""" & iif(SelectedState = "SD", "selected=""selected""", "") & ">South Dakota</option>" & vbCrlf
                    .Write "<option value=""TN""" & iif(SelectedState = "TN", "selected=""selected""", "") & ">Tennessee</option>" & vbCrlf
                    .Write "<option value=""TX""" & iif(SelectedState = "TX", "selected=""selected""", "") & ">Texas</option>" & vbCrlf
                    .Write "<option value=""UT""" & iif(SelectedState = "UT", "selected=""selected""", "") & ">Utah</option>" & vbCrlf
                    .Write "<option value=""VT""" & iif(SelectedState = "VT", "selected=""selected""", "") & ">Vermont</option>" & vbCrlf
                    .Write "<option value=""VA""" & iif(SelectedState = "VA", "selected=""selected""", "") & ">Virginia</option>" & vbCrlf
                    .Write "<option value=""WA""" & iif(SelectedState = "WA", "selected=""selected""", "") & ">Washington</option>" & vbCrlf
                    .Write "<option value=""WV""" & iif(SelectedState = "WV", "selected=""selected""", "") & ">West Virginia</option>" & vbCrlf
                    .Write "<option value=""WI""" & iif(SelectedState = "WI", "selected=""selected""", "") & ">Wisconsin</option>" & vbCrlf
                    .Write "<option value=""WY""" & iif(SelectedState = "WY", "selected=""selected""", "") & ">Wyoming</option>" & vbCrlf
            end select
            .Write "</select>" & vbCrlf
        end with
    end sub

    function DecodeId(Id)
        if IsNumeric(id) then
            dim step1

            step1 = clng(id) - ENCODE_SEED

            if step1 mod ENCODE_SEED = 0 then
                DecodeId = (CLng(Id) - ENCODE_SEED) / ENCODE_SEED
            else
                DecodeId = 0
            end if
        else
            DecodeId = 0
        end if
    end function

    'Function DoIIF(blnExpression, vTrueResult, vFalseResult)
    '    if blnExpression then
    '        doIIF = vTrueResult
    '    else
    '        doIIF = vFalseResult
    '    end if
    'End Function

    function EncodeId(Id)
        if IsNumeric(id) then
            EncodeId = (CLng(Id) * ENCODE_SEED) + ENCODE_SEED
        else
            EncodeId = 0
        end if
    end function

    function formatMyDate(mydate)
        dim newFormDate
        newFormDate = CDate(mydate)
        formatMyDate =  Month(newFormDate) & "/" & Day(newFormDate) & "/" & Year(newFormDate)
    end function 

    sub HandleMobileRedirect()
        if IsMobile() then
            select case Request.Cookies("likemobile")
                case "True"
                    Response.Redirect "http://www.gatsplat.com/mobile/index.asp"
                case "False"
                    ' do not redirect
                case else
                    Response.Redirect "http://www.gatsplat.com/mobile/mobiledetected.asp"
            end select
        end if
    end sub

    function IIf(expression, truePart, falsePart)
        if expression then
            IIf = truePart
        else
            IIf = falsePart
        end if
    end function

    function GetTimeForAdminCheckIn()
        dim myDate, checkInNow
        set myDate = new AvailableDate
        myDate.Load FormatDateTime(date(),vbshortdate)

        dim firstIntervalOfDay, lastIntervalOfDay, nextInterval, previousInterval
        firstIntervalOfDay = FormatDateTime(myDate.StartTime,vbShortTime)
        lastIntervalOfDay = FormatDateTime(DateAdd("n", (CInt(Settings(SETTING_BOOKING_INTERVAL)) * -1), myDate.EndTimeLongFormat), vbShortTime)

        checkInNow = FormatDateTime(Now(),vbShortTime)
        previousInterval = FormatDateTime(firstIntervalOfDay,vbShortTime)
        nextInterval = FormatDateTime(DateAdd("n", CInt(Settings(SETTING_BOOKING_INTERVAL)), GetTimeString(firstIntervalOfDay)), vbShortTime)

        If checkInNow >= lastIntervalOfDay Then
            checkInNow = lastIntervalOfDay
        Else
            If checkInNow <= firstIntervalOfDay Then
                checkInNow = firstIntervalOfDay
            Else
                Do While previousInterval <= lastIntervalOfDay 
                    If checkInNow >= FormatDateTime(previousInterval,vbShortTime) AND checkInNow <=  FormatDateTime(nextInterval,vbShortTime) Then
                        If checkInNow < FormatDateTime(DateAdd("n",6,previousInterval),vbShortTime) Then
                            checkInNow = previousInterval
                            Exit Do
                        Else
                            checkInNow = nextInterval
                            Exit Do
                        End If
                    End IF
                    previousInterval = nextInterval
                    nextInterval = FormatDateTime(DateAdd("n", CInt(Settings(SETTING_BOOKING_INTERVAL)), GetTimeString(previousInterval)), vbShortTime)
                Loop
            End If
        End If

        GetTimeForAdminCheckIn = checkInNow
    end function

    function GetGuid()
        dim typeLib
        set typeLib = Server.CreateObject("Scriptlet.TypeLib")

        GetGuid = Left(CStr(TypeLib.Guid), 38)

        Set TypeLib = Nothing
    end function

    function GetMilitaryTime(Time)
        GetMilitaryTime = left(Time, 5)
    end function

    function GetTimeString(Time)
        if len(time) >= 5 then 
            dim shortenedTime
            shortenedTime = left(Time, 5)

            dim timeParts
            timeParts = split(shortenedTime, ":")

            dim hourPart
            hourPart = cint(timeParts(0))

            dim meridian
            meridian = " AM"
        
            if hourPart = 0 then
                hourPart = 12
            elseif hourPart = 12 then
                meridian = " PM"
            elseif hourPart > 12 then
                hourPart = hourPart - 12
                meridian = " PM"
            end if

            GetTimeString = cstr(hourPart) & ":" & timeParts(1) & meridian
        else
            GetTimeString = Time
        end if
    end function

    function HTMLDecode(sText)
		Dim I

		sText = Replace(sText, "&amp;" , Chr(38))
		sText = Replace(sText, "&quot;", Chr(34))
		sText = Replace(sText, "&lt;"  , Chr(60))
		sText = Replace(sText, "&gt;"  , Chr(62))
		sText = Replace(sText, "&nbsp;", Chr(32))

		For I = 1 to 255
			sText = Replace(sText, "&#" & I & ";", Chr(I))
		Next

		HTMLDecode = sText
	end function

    Function IsMobile()
        Set Regex = New RegExp
        With Regex
            .Pattern = "(up.browser|up.link|mmp|symbian|smartphone|midp|wap|phone|windows ce|pda|mobile|mini|palm|ipad|silk)"
            .IgnoreCase = True
            .Global = True
        End With

        IsMobile = Regex.test(Request.ServerVariables("HTTP_USER_AGENT"))
    End Function

    function NullableReplace(Expression, Find, Replacement)
        if isnull(Expression) then
            NullableReplace = ""
        elseif IsNull(Replacement) then
            NullableReplace = Replace(Expression, Find, "", 1, -1, 1)
        else
            NullableReplace = Replace(Expression, Find, Replacement, 1, -1, 1)
        end if
    end function

    function PathCombine(Path, FileName)
        if len(Path) > 0 then
            dim endOfPath, beginOfFileName
            endOfPath = right(Path, 1)
            beginOfFileName = left(FileName, 1)

            if endOfPath <> "/" and endOfPath <> "\" and beginOfFileName <> "/" and beginOfFileName <> "\" then
                PathCombine = Path & "/" & FileName
            else
                PathCombine = Path & FileName
            end if
        else
            PatchCombine = FileName
        end if
    end function

    sub RedirectMobileToSpecificPage(PageUri)
        if IsMobile() then
            dim myCookie
            myCookie = Request.Cookies("likemobile")

            if myCookie = "" or myCookie = "True" then
                Response.Redirect(PageUri)
            end if
        end if
    end sub

    Function ReplaceIIFs(inString)
        Dim myStartIf, myEndIf, myResult, myExpression, myTrueString, myFalseString, myFoundAt, myEndAt
 
        myStartIf = InStr(1, inString, "{{IIF(")
        do while myStartIf > 0
            myEndIf = InStr(myStartIf, inString, ")}}")
            if myEndIf > 0 then
                myFoundAt = InStr(myStartIf, inString, ", """)
                if myFoundAt > 0 and myFoundAt < myEndIf then
                    myExpression = Mid(inString, myStartIf + 6, myFoundAt - (myStartIf + 6))
                    myEndAt = InStr(myFoundAt, inString, """, """)
                    if myEndAt > 0 and myEndAt < myEndIf then
                        myTrueString = Mid(inString, myFoundAt + 3, myEndAt - (myFoundAt + 3))
                        myFoundAt = InStr(myEndAt, inString, """)}}")
                        myResult = ""
                        if myFoundAt > 0 and myFoundAt < myEndIf then
                            myFalseString = Mid(inString, myEndAt + 4, myFoundAt - (myEndAt + 4))
                            on error resume next
                            myResult = iif(Eval(myExpression), myTrueString, myFalseString)
                            on error goto 0
                        end if
                        inString = Mid(inString, 1, myStartIf - 1) & myResult & Mid(inString, myEndIf + 3)
                    end if
                end if
                myStartIf = InStr(1, inString, "{{IIF(")
            else
                myStartIf = 0
            end if
        loop
           
        replaceIIFs = inString
    End Function
 
    function UsesPayments()
        dim depositFound : depositFound = false

        depositFound = Settings(SETTING_REGISTRATION_PAYMENTTYPE) <> PAYMENTTYPE_NONE

        if depositFound = false then
            dim customFields : set customFields = new RegistrationCustomFieldCollection
            customFields.LoadAll

            for i = 0 to customFields.Count - 1
                if cstr(customFields(i).PaymentTypeId) > PAYMENTTYPE_NONE  then
                    dim options : set options = new RegistrationCustomOptionCollection
                    options.LoadByFieldId customFields(i).RegistrationCustomFieldId

                    for j = 0 to options.Count - 1
                        if len(options(j).Value) > 0 then
                            depositFound = true
                            j = options.Count
                            i = customFields.Count
                        end if
                    next
                end if
            next
        end if

        UsesPayments = depositFound
    end function

	function URLDecode(sConvert)
		Dim aSplit
		Dim sOutput
		Dim I

		If IsNull(sConvert) or len(sConvert) = 0 Then
			URLDecode = ""
			Exit Function
		End If
			
		sOutput = REPLACE(sConvert, "+", " ")
			
		aSplit = Split(sOutput, "%")
			
		If IsArray(aSplit) Then
			sOutput = aSplit(0)
			For I = 0 to UBound(aSplit) - 1
				sOutput = sOutput & _
				Chr("&H" & Left(aSplit(i + 1), 2)) &_
				Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
			Next
		End If
			
		URLDecode = sOutput
	end function
%>