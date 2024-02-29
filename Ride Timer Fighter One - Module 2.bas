Public StopIt       As Boolean
Public ResetIt      As Boolean
Public LastTime
Sub RideTimerOne_Start()
    Dim StartTime, FinishTime, TotalTime, PauseTime
    StopIt = FALSE
    ResetIt = FALSE
    If Range("B8") = 0 Then
        StartTime = timer
        PauseTime = 0
        LastTime = 0
    Else
        StartTime = 0
        PauseTime = timer
    End If
    StartIt:
    DoEvents
    If StopIt = TRUE Then
        LastTime = TotalTime
        Exit Sub
    Else
        FinishTime = timer
        TotalTime = FinishTime - StartTime + LastTime - PauseTime
        TTime = TotalTime * 100
        HM = TTime Mod 100
        TTime = TTime \ 100
        hh = TTime \ 3600
        TTime = TTime Mod 3600
        MM = TTime \ 60
        SS = TTime Mod 60
        Range("B8").Value = Format(hh, "00") & ":" & Format(MM, "00") & ":" & Format(SS, "00") & "." & Format(HM, "00")
        If ResetIt = TRUE Then
            Range("B8") = Format(0, "00") & ":" & Format(0, "00") & ":" & Format(0, "00") & "." & Format(0, "00")
            LastTime = 0
            PauseTime = 0
            End
        End If
        GoTo StartIt
    End If
End Sub
Sub RideTimerOne_Stop()
    StopIt = TRUE
End Sub
Sub RideTimerOne_Reset()
    Range("B8").Value = Format(0, "00") & ":" & Format(0, "00") & ":" & Format(0, "00") & "." & Format(0, "00")
    LastTime = 0
    ResetIt = TRUE
End Sub
