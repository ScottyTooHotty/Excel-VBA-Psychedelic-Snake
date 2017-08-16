Attribute VB_Name = "Initialize"
Option Explicit

Public iScore As Long
Public lLength As Long
Public lStart As Long
Public lStep As Long
Public rHead As Range
Public rBody() As Range
Public rTail As Range
Public strMoveDirection As String

Sub StartGame()

    With Application
        .Cursor = xlWait
        .ScreenUpdating = False
        .OnKey "{ }"
    End With
    
    bStopTimer = False
    bStillGrowing = False
    lStart = 1
    lLength = 1
    iScore = -3
    strMoveDirection = "Right"
    ReDim rBody(lStart To lLength)
    
    With Sheet1
        .Cells(Rows.Count, Columns.Count).Select
        ActiveWindow.ScrollColumn = 1
        ActiveWindow.ScrollRow = 1
        .Protect UserInterfaceOnly:=True
        .Cells.Clear
        Call GameBoard(1)
        .EnableSelection = xlNoSelection
        
        Set rHead = _
            .Cells(Int(rGameGrid.Rows.Count / 2) + 1, Int(rGameGrid.Columns.Count / 2) + 1)
        For lStep = lStart To lLength Step 1
            Set rBody(lStep) = rHead.Offset(0, (lStep - 2) * 1)
        Next lStep
        Set rTail = rBody(lLength).Offset(0, -1)
        
        With rHead
            .BorderAround xlSolid, xlThick
            .Value = "O"
            .VerticalAlignment = xlCenter
            .HorizontalAlignment = xlCenter
            .Interior.ColorIndex = 1
            With .Font
                .Bold = True
                .ColorIndex = 2
            End With
        End With
        
        For lStep = lStart To lLength Step 1
            With rBody(lStep)
                .BorderAround LineStyle:=xlSolid, Weight:=xlThick
                .Interior.ColorIndex = RandomColour
            End With
        Next lStep
        With rTail
            .BorderAround xlSolid, xlThick
            .Interior.ColorIndex = RandomColour
        End With
    End With
    
    TimerID = SetTimer(0&, 0&, TimerSeconds, AddressOf TimerProc)
    
    Application.ScreenUpdating = True
        
End Sub
