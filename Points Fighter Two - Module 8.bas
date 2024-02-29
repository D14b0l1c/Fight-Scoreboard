Sub TakedownFighterTwo()
    Range("H2").Value = Range("H2").Value + 2
    
    Dim timeWS      As Worksheet
    Dim timeRange   As Range
    
    'Change "Fighter Two Logs" to whatever worksheet is your click log
    Set timeWS = ThisWorkbook.Sheets("Fighter Two Logs")
    
    'find the last cell in column C of your log
    Set timeRange = timeWS.Range("C" & timeWS.Rows.Count).End(xlUp).Offset(1)
    
    'Write which button was clicked in column C
    timeRange.Value = Application.Caller
    
    'Write the time in column D
    timeRange.Offset(, 1).Value = Now()
    
End Sub
Sub ReversalFighterTwo()
    Range("H3").Value = Range("H3").Value + 2
    
    Dim timeWS      As Worksheet
    Dim timeRange   As Range
    
    'Change "Fighter Two Logs" to whatever worksheet is your click log
    Set timeWS = ThisWorkbook.Sheets("Fighter Two Logs")
    
    'find the last cell in column E of your log
    Set timeRange = timeWS.Range("E" & timeWS.Rows.Count).End(xlUp).Offset(1)
    
    'Write which button was clicked in column E
    timeRange.Value = Application.Caller
    
    'Write the time in column F
    timeRange.Offset(, 1).Value = Now()
    
End Sub
Sub EscapeFighterTwo()
    Range("H4").Value = Range("H4").Value + 1
    
    Dim timeWS      As Worksheet
    Dim timeRange   As Range
    
    'Change "Fighter Two Logs" to whatever worksheet is your click log
    Set timeWS = ThisWorkbook.Sheets("Fighter Two Logs")
    
    'find the last cell in column G of your log
    Set timeRange = timeWS.Range("G" & timeWS.Rows.Count).End(xlUp).Offset(1)
    
    'Write which button was clicked in column G
    timeRange.Value = Application.Caller
    
    'Write the time in column H
    timeRange.Offset(, 1).Value = Now()
    
End Sub
Sub RunTimeFighterTwo()
    Range("H5").Value = Range("H5").Value + 1
End Sub
Sub PenaltyFighterTwo()
    Range("B6").Value = Range("B6").Value + 1
End Sub
Sub PenaltyXFighterTwo()
    Range("F16").Value = Range("F16").Value + 1
    
    Dim timeWS      As Worksheet
    Dim timeRange   As Range
    
    'Change "Fighter Two Logs" to whatever worksheet is your click log
    Set timeWS = ThisWorkbook.Sheets("Fighter Two Logs")
    
    'find the last cell in column I of your log
    Set timeRange = timeWS.Range("I" & timeWS.Rows.Count).End(xlUp).Offset(1)
    
    'Write which button was clicked in column I
    timeRange.Value = Application.Caller
    
    'Write the time in column J
    timeRange.Offset(, 1).Value = Now()
    
End Sub
