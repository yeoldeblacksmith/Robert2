<%
class PaymentType
    ' attributes
    private mn_Id, ms_Desc

    ' methods
    public sub Load()
        select case Id
            case 0
                Description = "None"
            case 1
                Description = "By Event"
            case 2
                Description = "By Player"
        end select
    end sub

    public sub LoadById(nId)
        Id = nId
        Load
    end sub

    ' properties
    public property get Id
        Id = mn_Id
    end property

    public property let Id(value)
        mn_Id = value
    end property

    public property get Description
        Description = ms_Desc
    end property

    public property let Description(value)
        ms_Desc = value
    end property
end class
%>