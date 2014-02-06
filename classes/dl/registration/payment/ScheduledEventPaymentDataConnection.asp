<%
class ScheduledEventPaymentDataConnection
    public function GetAllScheduledEventPaymentByEventId(EventId)
        dim oCmd, oRs, oCon
        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetAllScheduledEventPaymentByEventId"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)

        oCon.GetRecordSet oCmd, oRs

        dim results
        if oRs.EOF then
            results = Array()
        else
            results = oRs.GetRows()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetAllScheduledEventPaymentByEventId = results
    end function

    public function GetLastScheduledEventPaymentByEventId(EventId)
        dim oCmd, oRs, oCon
        set oCmd = Server.CreateObject("ADODB.Command")
        set oRs = Server.CreateObject("ADODB.Recordset")
        set oCon = new DataConnection

        oCmd.CommandText = "GetLastScheduledEventPaymentByEventId"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)

        oCon.GetRecordSet oCmd, oRs

        dim results
        if oRs.EOF then
            results = Array()
        else
            results = oRs.GetRows()
        end if

        oRs.Close
        oCon.CloseConnection

        set oRs = nothing
        set oCmd = nothing

        GetLastScheduledEventPaymentByEventId = results
    end function

    public sub SaveScheduledEventPayment(EventId, StatusDateTime, PaymentAmount, PaymentStatus, PayPalTransactionId)
        dim oCmd, oCon
        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "SaveScheduledEventPayment"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)

        if IsNull(StatusDateTime) = false and len(StatusDateTime) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@StatusDateTime", adDBDate, adParamInput, 8, StatusDateTime)
        else 
            oCmd.Parameters.Append oCmd.CreateParameter("@StatusDateTime", adDBDate, adParamInput, 8, null)
        end if
    
        if IsNull(PaymentAmount) = false and len(PaymentAmount) > 0 then
            dim parmPaymentAmount : set parmPaymentAmount = oCmd.CreateParameter("@PaymentAmount", adDecimal, adParamInput, 9, PaymentAmount)
            parmPaymentAmount.NumericScale = 2
            parmPaymentAmount.Precision = 9
            oCmd.Parameters.Append parmPaymentAmount
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@PaymentAmount", adDecimal, adParamInput, 9, null)
        end if
    
        if IsNull(PaymentStatus) = false and len(PaymentStatus) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@PaymentStatus", adVarChar, adParamInput, 50, PaymentStatus)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@PaymentStatus", adVarChar, adParamInput, 50, null)
        end if

        if IsNull(PayPalTransactionId) = false and len(PayPalTransactionId) > 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@PayPalTransactionId", adVarChar, adParamInput, 128, PayPalTransactionId)    
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@PayPalTransactionId", adVarChar, adParamInput, 128, null)    
        end if

        oCon.ExecuteCommand oCmd
        oCon.CloseConnection

        set oCmd = nothing
    end sub
end class
%>