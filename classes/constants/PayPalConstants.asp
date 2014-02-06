<%
    const PAYPAL_IPN_ADDRESS = "https://www.paypal.com/cgi-bin/webscr"

    const PAYPAL_VARIABLE_BUSINESS = "business"
    const PAYPAL_VARIABLE_COMMENTS = "memo"
    const PAYPAL_VARIABLE_ITEMDESCRIPTION = "item_name"
    const PAYPAL_VARIABLE_ITEMLINENUMBER = "item_number"
    const PAYPAL_VARIABLE_ITEMPRICE = "amount"
    const PAYPAL_VARIABLE_ITEMQTY = "quantity"
    const PAYPAL_VARIABLE_EVENTID = "custom"
    const PAYPAL_VARIABLE_LINECOUNT = "num_cart_items"
    const PAYPAL_VARIABLE_PAYMENTAMOUNT = "mc_gross"
    const PAYPAL_VARIABLE_PAYMENTDATE = "payment_date"
    const PAYPAL_VARIABLE_PAYMENTSTATUS = "payment_status"
    const PAYPAL_VARIABLE_TRANSACTIONID = "txn_id"

    const PAYPAL_PAYMENTSTATUS_PAID = "Completed"
    const PAYPAL_PAYMENTSTATUS_REQUESTED = "Requested"

    const PAYPAL_TRANSID_MANUAL = "MARKED_BY_ADMIN"
%>