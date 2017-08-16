Attribute VB_Name = "Timing"
Option Explicit

Public Declare PtrSafe Function SetTimer Lib "user32" ( _
    ByVal HWnd As Long, _
    ByVal nIDEvent As Long, _
    ByVal uElapse As Long, _
    ByVal lpTimerFunc As Long) As Long

Public Declare PtrSafe Function KillTimer Lib "user32" ( _
    ByVal HWnd As Long, _
    ByVal nIDEvent As Long) As Long

Public bStopTimer As Boolean
Public TimerID As Long
Public TimerSeconds As Single

Sub EndTimer()
    KillTimer 0&, TimerID
    bStopTimer = True
    With Application
        .Cursor = xlDefault
        .OnKey "{ }", "StartGame"
    End With
    Sheet1.EnableSelection = xlNoRestrictions
    Sheet1.Unprotect
    Sheet1.Range("A1").Select
End Sub

Public Sub TimerProc(ByVal HWnd As Long, ByVal uMsg As Long, _
        ByVal nIDEvent As Long, ByVal dwTimer As Long)

    Call Move(strMoveDirection)
    
End Sub


