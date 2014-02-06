<%
class EmailConnection
    ' attributes
    private mo_emailClient, ms_SenderName, ms_SenderAddress, _
            ms_ServerAddress, ms_InfoAddress

    ' ctor
    public sub Class_Initialize()
        set mo_emailClient = Server.CreateObject("SMTPsvg.Mailer")
    end sub

    ' methods
    public function Send(ToName, ToAddress, Subject, Body)
        on error resume next
        with mo_emailClient
            .FromName = Settings(SETTING_FROMNAME)
            .FromAddress = Settings(SETTING_FROMADDRESS)
            .RemoteHost = Settings(SETTING_EMAILSERVER)
            .AddRecipient ToName, ToAddress
            .Subject = Subject
            .BodyText = Body
            .ContentType = "text/html"

            Send = .SendMail()
        end with
    end function

    public function SendToReservation(FromName, FromAddress, Subject, Body)
        on error resume next
        with mo_emailClient
            .FromName = FromName
            .FromAddress = FromAddress
            .RemoteHost = ServerAddress
            .AddRecipient SenderName, InfoAddress
            .Subject = Subject
            .BodyText = Body
            .ContentType = "text/html"

            SendToReservation = .SendMail()
        end with
    end function

    public property get InfoAddress
        InfoAddress = ms_InfoAddress
    end property

    public property let InfoAddress(value)
        ms_InfoAddress = value
    end property

    public property get SenderAddress
        SenderAddress = ms_SenderAddress
    end property
        
    public property let SenderAddress(value)
        ms_SenderAddress = value
    end property

    public property get SenderName
        SenderName = ms_SenderName
    end property

    public property let SenderName(value)
        ms_SenderName = value
    end property

    public property get ServerAddress
        ServerAddress = ms_ServerAddress
    end property

    public property let ServerAddress(value)
        ms_ServerAddress = value
    end property
end class
%>