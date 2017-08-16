Attribute VB_Name = "Movement"
Option Explicit

Sub Move(strDirection As String)
    
    On Error GoTo Kill
    
    With Application
        .ScreenUpdating = False
        .OnKey "{UP}", "'KeyPress ""Up""'"
        .OnKey "{DOWN}", "'KeyPress ""Down""'"
        .OnKey "{LEFT}", "'KeyPress ""Left""'"
        .OnKey "{RIGHT}", "'KeyPress ""Right""'"
    End With

    rTail.ClearFormats
    For lStep = lLength To lStart Step -1
        rBody(lStep).ClearFormats
    Next lStep
    With rHead
        .ClearFormats
        .Value = ""
    End With

    Select Case strDirection
        Case "Up"
            If Intersect(rHead.Offset(-1, 0), rGameGrid) Is Nothing Then
                If rHead.Offset(rGameGrid.Rows.Count - 1, 0).Value = "X" Then
                    'If bStillGrowing = True Then
                    '    Call Eat(False)
                    'Else
                        Call Eat(True)
                    'End If
                ElseIf bStillGrowing = True Then
                    Call Eat(False)
                Else
                    Set rTail = rBody(lLength)
                End If
            Else
                If rHead.Offset(-1, 0).Value = "X" Then
                    'If bStillGrowing = True Then
                    '    Call Eat(False)
                    'Else
                        Call Eat(True)
                    'End If
                ElseIf bStillGrowing = True Then
                    Call Eat(False)
                Else
                    Set rTail = rBody(lLength)
                End If
            End If
        Case "Down"
            If Intersect(rHead.Offset(1, 0), rGameGrid) Is Nothing Then
                If rHead.Offset((rGameGrid.Rows.Count - 1) * -1, 0).Value = "X" Then
                    'If bStillGrowing = True Then
                    '    Call Eat(False)
                    'Else
                        Call Eat(True)
                    'End If
                ElseIf bStillGrowing = True Then
                    Call Eat(False)
                Else
                    Set rTail = rBody(lLength)
                End If
            Else
                If rHead.Offset(1, 0).Value = "X" Then
                    'If bStillGrowing = True Then
                    '    Call Eat(False)
                    'Else
                        Call Eat(True)
                    'End If
                ElseIf bStillGrowing = True Then
                    Call Eat(False)
                Else
                    Set rTail = rBody(lLength)
                End If
            End If
        Case "Left"
            If Intersect(rHead.Offset(0, -1), rGameGrid) Is Nothing Then
                If rHead.Offset(0, rGameGrid.Rows.Count - 1).Value = "X" Then
                    'If bStillGrowing = True Then
                    '    Call Eat(False)
                    'Else
                        Call Eat(True)
                    'End If
                ElseIf bStillGrowing = True Then
                    Call Eat(False)
                Else
                    Set rTail = rBody(lLength)
                End If
            Else
                If rHead.Offset(0, -1).Value = "X" Then
                    'If bStillGrowing = True Then
                    '    Call Eat(False)
                    'Else
                        Call Eat(True)
                    'End If
                ElseIf bStillGrowing = True Then
                    Call Eat(False)
                Else
                    Set rTail = rBody(lLength)
                End If
            End If
        Case "Right"
            If Intersect(rHead.Offset(0, 1), rGameGrid) Is Nothing Then
                If rHead.Offset(0, (rGameGrid.Rows.Count - 1) * -1).Value = "X" Then
                    'If bStillGrowing = True Then
                    '    Call Eat(False)
                    'Else
                        Call Eat(True)
                    'End If
                ElseIf bStillGrowing = True Then
                    Call Eat(False)
                Else
                    Set rTail = rBody(lLength)
                End If
            Else
                If rHead.Offset(0, 1).Value = "X" Then
                    'If bStillGrowing = True Then
                    '    Call Eat(False)
                    'Else
                        Call Eat(True)
                    'End If
                ElseIf bStillGrowing = True Then
                    Call Eat(False)
                Else
                    Set rTail = rBody(lLength)
                End If
            End If
        Case Else
    End Select

    For lStep = lLength To lStart Step -1
        If lStep = lStart Then
            Set rBody(lStep) = rHead
        Else
            Set rBody(lStep) = rBody(lStep - 1)
        End If
    Next lStep
    
    Select Case strDirection
        Case "Up"
            If Intersect(rHead.Offset(-1, 0), rGameGrid) Is Nothing Then
                Set rHead = rHead.Offset(rGameGrid.Rows.Count - 1, 0)
            Else
                Set rHead = rHead.Offset(-1, 0)
            End If
        Case "Down"
            If Intersect(rHead.Offset(1, 0), rGameGrid) Is Nothing Then
                Set rHead = rHead.Offset((rGameGrid.Rows.Count - 1) * -1, 0)
            Else
                Set rHead = rHead.Offset(1, 0)
            End If
        Case "Left"
            If Intersect(rHead.Offset(0, -1), rGameGrid) Is Nothing Then
                Set rHead = rHead.Offset(0, rGameGrid.Columns.Count - 1)
            Else
                Set rHead = rHead.Offset(0, -1)
            End If
        Case "Right"
            If Intersect(rHead.Offset(0, 1), rGameGrid) Is Nothing Then
                Set rHead = rHead.Offset(0, (rGameGrid.Columns.Count - 1) * -1)
            Else
                Set rHead = rHead.Offset(0, 1)
            End If
        Case Else
    End Select

    rGameGrid.BorderAround LineStyle:=xlContinuous, Weight:=xlThick, ColorIndex:=RandomColour
    
    With rHead
        .BorderAround LineStyle:=xlSolid, Weight:=xlThick
        .Value = "O"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.ColorIndex = 1
        With .Font
            .Bold = True
            .ColorIndex = 2
        End With
    End With
    
    For lStep = lLength To lStart Step -1
        With rBody(lStep)
            .BorderAround LineStyle:=xlSolid, Weight:=xlThick
            .Interior.ColorIndex = RandomColour
        End With
    Next lStep
    With rTail
        .BorderAround LineStyle:=xlSolid, Weight:=xlThick
        .Interior.ColorIndex = RandomColour
    End With
    
    iApples = WorksheetFunction.CountIf(rGameGrid, "X")
    If iApples < iTotalApples Then
        Call TopUp
    End If

    If bStopTimer = True Then
        Call EndTimer
    End If
    
    For lStep = lLength To lStart Step -1
        If rHead.Address = rBody(lStep).Address Then
            GoTo Kill
        End If
    Next lStep
    If rHead.Address = rTail.Address Then
        GoTo Kill
    End If
    
    Application.ScreenUpdating = True
    
    Exit Sub
    
Kill:
    Call EndTimer
    
    Application.ScreenUpdating = True

End Sub

Sub KeyPress(strKey As String)

    Select Case strKey
        Case "Up"
            If strMoveDirection <> "Down" Then
                strMoveDirection = "Up"
            End If
        Case "Down"
            If strMoveDirection <> "Up" Then
                strMoveDirection = "Down"
            End If
        Case "Right"
            If strMoveDirection <> "Left" Then
                strMoveDirection = "Right"
            End If
        Case "Left"
            If strMoveDirection <> "Right" Then
                strMoveDirection = "Left"
            End If
        Case Else
    End Select
    
    With Application
        .OnKey "{UP}"
        .OnKey "{DOWN}"
        .OnKey "{LEFT}"
        .OnKey "{RIGHT}"
    End With

End Sub
