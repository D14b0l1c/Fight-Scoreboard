Public interval     As Date
Sub start_overtimer_fighter_one()
    
    interval = Now + TimeValue("00:00:01")
    
    If Range("B15").Value = 0 Then Exit Sub
    
    Range("B15") = Range("B15") - TimeValue("00:00:01")
    
    Application.OnTime interval, "start_overtimer_fighter_one"
    
End Sub
Sub stop_overtimer_fighter_one()
    
    Application.OnTime EarliestTime:=interval, Procedure:="start_overtimer_fighter_one", Schedule:=False
    
End Sub
Sub reset_overtimer_fighter_one()
    Range("B15").Value = "00:00:30"
End Sub
