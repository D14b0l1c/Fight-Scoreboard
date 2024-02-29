Public interval     As Date
Sub start_overtimer_one()
    
    interval = Now + TimeValue("00:00:01")
    
    If Range("E9").Value = 0 Then Exit Sub
    
    Range("E9") = Range("E9") - TimeValue("00:00:01")
    
    Application.OnTime interval, "start_overtimer_one"
    
End Sub
Sub stop_overtimer_one()
    
    Application.OnTime EarliestTime:=interval, Procedure:="start_overtimer_one", Schedule:=False
    
End Sub
Sub reset_overtimer_one()
    Range("E9").Value = "00:02:00"
End Sub
