Sub TakedownFighterOne()
    Range("B2").Value = Range("B2").Value + 2
    
    Dim timeWS      As Worksheet
    Dim timeRange   As Range
    
    'Change "Fighter One Logs" to whatever worksheet is your click log
    Set timeWS = ThisWorkbook.Sheets("Fighter One Logs")
    
    'find the last cell in column C of your log
    Set timeRange = timeWS.Range("C" & timeWS.Rows.Count).End(xlUp).Offset(1)
    
    'Write which button was clicked in column C
    timeRange.Value = Application.Caller
    
    'Write the time in column D
    timeRange.Offset(, 1).Value = Now()
    
End Sub
Sub ReversalFighterOne()
    Range("B3").Value = Range("B3").Value + 2
    
    Dim timeWS      As Worksheet
    Dim timeRange   As Range
    
    'Change "Fighter One Logs" to whatever worksheet is your click log
    Set timeWS = ThisWorkbook.Sheets("Fighter One Logs")
    
    'find the last cell in column E of your log
    Set timeRange = timeWS.Range("E" & timeWS.Rows.Count).End(xlUp).Offset(1)
    
    'Write which button was clicked in column E
    timeRange.Value = Application.Caller
    
    'Write the time in column F
    timeRange.Offset(, 1).Value = Now()
    
End Sub
Sub EscapeFighterOne()
    Range("B4").Value = Range("B4").Value + 1
    
    Dim timeWS      As Worksheet
    Dim timeRange   As Range
    
    'Change "Fighter One Logs" to whatever worksheet is your click log
    Set timeWS = ThisWorkbook.Sheets("Fighter One Logs")
    
    'find the last cell in column G of your log
    Set timeRange = timeWS.Range("G" & timeWS.Rows.Count).End(xlUp).Offset(1)
    
    'Write which button was clicked in column G
    timeRange.Value = Application.Caller
    
    'Write the time in column H
    timeRange.Offset(, 1).Value = Now()
    
End Sub
Sub RunTimeFighterOne()
    Range("B5").Value = Range("B5").Value + 1
End Sub
Sub PenaltyFighterOne()
    Range("H6").Value = Range("H6").Value + 1
End Sub
Sub PenaltyXFighterOne()
    Range("D16").Value = Range("D16").Value + 1
    
    Dim timeWS      As Worksheet
    Dim timeRange   As Range
    
    'Change "Fighter One Logs" to whatever worksheet is your click log
    Set timeWS = ThisWorkbook.Sheets("Fighter One Logs")
    
    'find the last cell in column I of your log
    Set timeRange = timeWS.Range("I" & timeWS.Rows.Count).End(xlUp).Offset(1)
    
    'Write which button was clicked in column I
    timeRange.Value = Application.Caller
    
    'Write the time in column J
    timeRange.Offset(, 1).Value = Now()
    
End Sub
