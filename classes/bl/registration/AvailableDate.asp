<% 
class AvailableDate
    private ms_Date, ms_StartTime, ms_EndTime, mb_Loaded

    'ctor
    public sub Class_Initialize()
        mb_Loaded = false
    end sub

    'methods
    public sub Delete(oDate)
        dim oDC
        set oDC = new AvailableDateDataConnection

        oDC.DeleteAvailableDate(oDate)
    end sub

    public function Exists(oDate)
        if mb_Loaded then
            Exists = true
        else
            dim oDC
            set oDC = new AvailableDateDataConnection

            Exists = not IsEmpty(oDC.GetAvailableDate(oDate))
        end if
    end function

    public sub Load(oDate)
        dim oDC
        set oDC = new AvailableDateDataConnection

        dim results
        results = oDC.GetAvailableDate(oDate)

        if IsEmpty(results) = false then
            SelectedDate = results(AVAILABLEDATE_INDEX_DATE)
            StartTime = results(AVAILABLEDATE_INDEX_STARTTIME)
            EndTime = results(AVAILABLEDATE_INDEX_ENDTIME)

            mb_loaded = true
        else
            SelectedDate = oDate

            select case Weekday(oDate)
                case 1
                    StartTime = "10:00"
                    EndTime = "17:00"
                case 2, 3, 4, 5
                    StartTime = "12:00"
                    EndTime = "18:00"
                case 6
                    StartTime = "14:00"
                    EndTime = "22:00"
                case 7
                    StartTime = "10:00"
                    EndTime = "20:00"
            end select

            mb_loaded = false
        end if
    end sub

    public sub Save(sDate, sStartTime, sEndTime)
        dim oDC
        set oDC = new AvailableDateDataConnection

        oDC.SaveAvailableDate sDate, sStartTime, sEndTime
    end sub

    public function ToDate()
        ToDate = CDate(SelectedDate)
    end function

    ' properties
    public property Get SelectedDate
        SelectedDate = ms_date
    end property

    public property Let SelectedDate(value)
        ms_date = value
    end property

    public property get StartTime
        StartTime = ms_StartTime
    end property

    public property let StartTime(value)
        ms_StartTime = value
    end property

    public property get StartTimeLongFormat
        StartTimeLongFormat = GetTimeString(ms_StartTime)
    end property

    public property get StartTimeShortFormat
        StartTimeShortFormat = GetMilitaryTime(ms_StartTime)
    end property

    public property get EndTime
        EndTime = ms_EndTime
    end property

    public property let EndTime(value)
        ms_EndTime = value
    end property

    public property get EndTimeLongFormat
        EndTimeLongFormat = GetTimeString(ms_EndTime)
    end property

    public property get EndTimeShortFormat
        EndTimeShortFormat = GetMilitaryTime(ms_EndTime)
    end property

    public property get Loaded
        Loaded = mb_loaded
    end property
end class
%>