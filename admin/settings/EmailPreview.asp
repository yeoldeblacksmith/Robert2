<!--#include file="../../classes/IncludeList.asp" -->
<%

	If Request.ServerVariables("REQUEST_METHOD") = "POST" then

		dim previewType, emailSubject, emailTemplate, emailPreview

		emailSubject = Request.Form("emailSubject")
		emailContent = URLDecode(Request.Form("emailContent"))

        select case Request.Form("previewType")
            case SETTING_EMAILTEMPLATE_CANCELLATION
			    set emailPreview = new CancellationEmail
            case SETTING_EMAILTEMPLATE_CONFIRMATION
    			set emailPreview = new ConfirmationEmail
            case SETTING_EMAILTEMPLATE_PAYMENTFAILED
                set emailPreview = new PaymentFailedEmail
            case SETTING_EMAILTEMPLATE_PAYMENTPENDING
                set emailPreview = new PaymentPendingEmail
            case SETTING_EMAILTEMPLATE_REMINDER
    			set emailPreview = new ReminderEmail
            case SETTING_EMAILTEMPLATE_WAIVER
    			set emailPreview = new WaiverEmail
            case SETTING_EMAILTEMPLATE_WAIVERLOOKUP
    			set emailPreview = new LookupEmail
        end select

	end if
%>
<html>
	<head>
		<title>Email Preview</title>
		<style type="text/css">
			body {
				background: #2D95E2;
				background: -webkit-gradient(linear, left top, left bottom, from(#2D95E2), to(#BAE7FF));
				background: -webkit-linear-gradient(#2D95E2, #BAE7FF);
				background: -moz-linear-gradient(top, #2D95E2, #BAE7FF);
				background: -ms-linear-gradient(#2D95E2, #BAE7FF);
				background: -o-linear-gradient(#2D95E2, #BAE7FF);
				background: linear-gradient(#2D95E2, #BAE7FF);
			}
		</style>
	</head>
	<body >
		<%
			if (IsObject(emailPreview)) then
				Response.Write(emailPreview.GetPreview(emailSubject, emailContent))
			end if
		%>
	</body>
</html>

