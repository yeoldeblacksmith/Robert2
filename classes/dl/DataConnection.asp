<% 
Class DataConnection
    ' attributes
    private m_connString 
    private m_con

    ' ctor
    public sub Class_Initialize()
        'Staging
        'm_connString = "Provider=SQLOLEDB.1;Password=T%dp^95-17[!;Persist Security Info=True;User ID=vantuserprod;Initial Catalog=vantorastaging;Data Source=64.251.193.8"

        'TestCustomer
        m_connString = "Provider=SQLOLEDB.1;Password=T%dp^95-17[!;Persist Security Info=True;User ID=vantuserprod;Initial Catalog=vantoraTestCustomer;Data Source=64.251.193.8"
    
        'Demo
        'm_connString = "Provider=SQLOLEDB.1;Password=[dem0ar0tnaV];Persist Security Info=True;User ID=vantdemo;Initial Catalog=vantorademo;Data Source=64.251.193.8"


        set m_con = Server.CreateObject("ADODB.Connection")
    end sub

    ' dtor
    public sub Class_Terminate()
        if m_con.State = adStateOpen then m_con.Close
    end sub

    ' methods
    public sub CloseConnection()
        if m_con.State = adStateOpen then m_con.Close
    end sub

    public sub ExecuteCommand(oCmd)
        if m_con.State <> adStateOpen then m_con.Open m_connString
        oCmd.ActiveConnection = m_con
        oCmd.Execute
    end sub

    public sub GetRecordSet(oCmd, oRs)
        if m_con.State <> adStateOpen then m_con.Open m_connString
        oCmd.ActiveConnection = m_con

        set oRs = Server.CreateObject("ADODB.Recordset")
        oRs.CursorLocation = adUseClient
        oRs.Open oCmd, , adOpenForwardOnly, adLockReadOnly
    end sub

end class
%>
