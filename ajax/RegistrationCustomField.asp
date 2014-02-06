<!--#include file="../classes/IncludeList.asp" -->
<%
    const AJAXACTION_RCF_CHANGESEQUENCE = "0"

    select case request.QueryString(QUERYSTRING_VAR_ACTION)
        case AJAXACTION_RCF_CHANGESEQUENCE
            ChangeSequence request.QueryString(QUERYSTRING_VAR_ID), _
                           request.QueryString(QUERYSTRING_VAR_SEQUENCE)
    end select


    private sub ChangeSequence(RegistrationCustomFieldId, Sequence)
        dim myField : set myField = new RegistrationCustomField
        myField.LoadById RegistrationCustomFieldId

        myField.Sequence = Sequence
        myField.Save
    end sub
%>