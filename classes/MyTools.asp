<%

Dim StopWatch(19), TotalElapsed
Dim ElapsedLap(19)

sub StartTimer(x)
   StopWatch(x) = timer
end sub


function StopTimer(x)
   EndTime = Timer

   'Watch for the midnight wraparound...
   if EndTime < StopWatch(x) then
     EndTime = EndTime + (86400)
   end if

   StopTimer = EndTime - StopWatch(x)
   StopWatch(x) = Timer    
end function

'StartTimer x

'ElapsedLap(x) = StopTimer(x)
'TotalElapsed = TotalElapsed + Elapsed(x)
'Response.Write "Elapsed time was: " & Elapsed

    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
%>