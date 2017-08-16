Attribute VB_Name = "Level"
Option Explicit

Public rGameGrid As Range

Sub GameBoard(lLevel As Long)
    
    With Sheet1
        Set rGameGrid = .Range(.Cells(2, 2), .Cells(18, 36))
    End With

    Select Case lLevel
        Case 1
            iApples = 0
            TimerSeconds = 125
            rGameGrid.BorderAround LineStyle:=xlContinuous, Weight:=xlThick, _
                ColorIndex:=RandomColour
            Call TopUp
        Case 2
            TimerSeconds = Int(TimerSeconds / 1.1)
            KillTimer 0&, TimerID
            TimerID = SetTimer(0&, 0&, TimerSeconds, AddressOf TimerProc)
            'rGameGrid.BorderAround LineStyle:=xlContinuous, Weight:=xlThick, _
                ColorIndex:=RandomColour
            'Call TopUp
        Case Else
    End Select
    
End Sub
