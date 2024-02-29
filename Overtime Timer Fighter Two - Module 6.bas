Public interval     As Date
Sub start_overtimer_fighter_two()
    
    interval = Now + TimeValue("00:00:01")
    
    If Range("H15").Value = 0 Then Exit Sub
    
    Range("H15") = Range("H15") - TimeValue("00:00:01")
    
    Application.OnTime interval, "start_overtimer_fighter_two"
    
End Sub
Sub stop_overtimer_fighter_two()
    
    Application.OnTime EarliestTime:=interval, Procedure:="start_overtimer_fighter_two", Schedule:=False
    
End Sub
Sub reset_overtimer_fighter_two()
    Range("H15").Value = "00:00:30"
End Sub
