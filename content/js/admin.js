/*
*	PreviewForm
*	Pops up a preview of the email in a new window
*/
function previewEmail(previewType, emailSubject, emailContent) {
	window.open('#', 'previewEmail', 'toolbar=no,width=670,height=720,location=no,status=no,scrollbars=yes');
	$form = $('<form id="hiddenPreview" action="EmailPreview.asp" method="POST" target="previewEmail"></form>');
	$form.append("<input type='hidden' name='previewType' value ='" + previewType + "' />");
	$form.append("<input type='hidden' name='emailSubject' value='" + emailSubject + "' />");
	$form.append('<input type="hidden" name="emailContent" value="' + escape(emailContent) + '" />');
	$('body').append($form);

	$form.submit();

	$form.remove();
}