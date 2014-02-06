<%
class OrderDetailDataConnection
    sub DeleteAllOrderDetailsByEventId(EventId)
        dim oCmd, oCon

        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "DeleteAllOrderDetailsByEventId"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)

        oCon.ExecuteCommand oCmd

        oCon.CloseConnection
        set oCmd = nothing
    end sub

    function GetAllOrderDetailsByEventId(EventId)
        dim oCmd, oRs, oCon
        set oCmd = Server.CreateObject("ADODB.Command")

        oCmd.CommandText = "GetAllOrderDetailsByEventId"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)

        set oCon = new DataConnection
        set oRs = Server.CreateObject("ADODB.Recordset")
        oCon.GetRecordset oCmd, oRs

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

        GetAllOrderDetailsByEventId = results
    end function

    function GetOrderDetail(EventId, LineNumber)
        dim oCmd, oRs, oCon
        set oCmd = Server.CreateObject("ADODB.Command")

        oCmd.CommandText = "GetOrderDetail"
        oCmd.CommandType = adCmdStoredProc
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)
        oCmd.Parameters.Append oCmd.CreateParameter("@LineNumber", adSmallInt, adParamInput, 4, LineNumber)

        set oCon = new DataConnection
        set oRs = Server.CreateObject("ADODB.Recordset")
        oCon.GetRecordset oCmd, oRs

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

        GetOrderDetail = results
    end function

    sub SaveOrderDetail(EventId, LineNumber, ItemNumber, _
                        ItemDescription, Quantity, Price, _
                        CreateDateTime, PaymentTransaction, PaymentDateTime, _
                        PaymentAmount, IsAdjustment)
        dim oCmd, oCon
        set oCmd = Server.CreateObject("ADODB.Command")
        set oCon = new DataConnection

        oCmd.CommandText = "SaveOrderDetail"
        oCmd.CommandType = adCmdStoredProc

        ' required parameters
        oCmd.Parameters.Append oCmd.CreateParameter("@EventId", adInteger, adParamInput, 4, EventId)
        oCmd.Parameters.Append oCmd.CreateParameter("@LineNumber", adSmallInt, adParamInput, 2, LineNumber)
        oCmd.Parameters.Append oCmd.CreateParameter("@ItemNumber", adInteger, adParamInput, 4, ItemNumber)
        oCmd.Parameters.Append oCmd.CreateParameter("@ItemDescription", adVarChar, adParamInput, 128, ItemDescription)
        oCmd.Parameters.Append oCmd.CreateParameter("@Quantity", adSmallInt, adParamInput, 2, Quantity)

        dim parmPrice : set parmPrice = oCmd.CreateParameter("@Price", adDecimal, adParamInput, 9, Price)
        parmPrice.NumericScale = 2
        parmPrice.Precision = 9
        oCmd.Parameters.Append parmPrice

        oCmd.Parameters.Append oCmd.CreateParameter("@CreateDateTime", adDate, adParamInput, 8, CreateDateTime)

        ' optional parameters
        if IsNull(PaymentTransaction) or len(PaymentTransaction) = 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@PaymentTransaction", adVarChar, adParamInput, 128, null)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@PaymentTransaction", adVarChar, adParamInput, 128, PaymentTransaction)
        end if

        if IsNull(PaymentDateTime) or len(PaymentDateTime) = 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@PaymentDateTime", adDate, adParamInput, 8, null)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@PaymentDateTime", adDate, adParamInput, 8, PaymentDateTime)
        end if

        if IsNull(PaymentAmount) or len(PaymentAmount) = 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@PaymentAmount", adDecimal, adParamInput, 9, Price)
        else
            dim parmPaymentAmount : set parmPaymentAmount = oCmd.CreateParameter("@PaymentAmount", adDecimal, adParamInput, 9, PaymentAmount)
            parmPaymentAmount.NumericScale = 2
            parmPaymentAmount.Precision = 9
            oCmd.Parameters.Append parmPaymentAmount
        end if

        if IsNull(IsAdjustment) or len(IsAdjustment) = 0 then
            oCmd.Parameters.Append oCmd.CreateParameter("@IsAdjustment", adBoolean, adParamInput, 1, null)
        else
            oCmd.Parameters.Append oCmd.CreateParameter("@IsAdjustment", adBoolean, adParamInput, 1, IsAdjustment)
        end if

        oCon.ExecuteCommand oCmd
        oCon.CloseConnection

        set oCmd = nothing
    end sub
end class
%>