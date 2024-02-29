Public interval     As Date
Sub start_timer()
    
    interval = Now + TimeValue("00:00:01")
    
    If Range("E2").Value = 0 Then Exit Sub
    
    Range("E2") = Range("E2") - TimeValue("00:00:01")
    
    Application.OnTime interval, "start_timer"
    
    Dim timeWS      As Worksheet
    Dim timeRange   As Range
    
    'Change "Sheet1" to whatever worksheet is your click log
    Set FighterOneWS = ThisWorkbook.Sheets("Fighter One Logs")
    Set FighterTwoWS = ThisWorkbook.Sheets("Fighter Two Logs")
    
    'find the last cell in column A of your log
    Set timeRangeOne = FighterOneWS.Range("A" & FighterOneWS.Rows.Count).End(xlUp).Offset(1)
    Set timeRangeTwo = FighterTwoWS.Range("A" & FighterTwoWS.Rows.Count).End(xlUp).Offset(1)
    
    'Write which button was clicked in column A
    timeRangeOne.Value = Application.Caller
    timeRangeTwo.Value = Application.Caller
    
    'Write the time in column B
    timeRangeOne.Offset(, 1).Value = Now()
    timeRangeTwo.Offset(, 1).Value = Now()
    
End Sub
Sub stop_timer()
    
    Application.OnTime EarliestTime:=interval, Procedure:="start_timer", Schedule:=False
    
End Sub
Sub reset_timer()
    Range("E2").Value = "00:02:00"
End Sub
