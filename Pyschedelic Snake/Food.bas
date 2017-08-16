Attribute VB_Name = "Food"
Option Explicit

Public iAcross As Integer
Public iApples As Integer
Public iDown As Integer
Public iTotalApples As Integer

Sub TopUp()

    Randomize
    
    iTotalApples = 3
    
    Do
        iDown = Int(rGameGrid.Rows.Count * Rnd) + 1
        iAcross = Int(rGameGrid.Columns.Count * Rnd) + 1
        If Sheet1.Cells(iDown + 1, iAcross + 1).Value = "" And _
            CellHasBorder(Sheet1.Cells(iDown + 1, iAcross + 1)) = False Then
            Sheet1.Cells(iDown + 1, iAcross + 1).Value = "X"
            iApples = iApples + 1
            iScore = iScore + 1
            With Sheet1.Range("AN3")
                .Value = "Score"
                .Font.Bold = True
            End With
            With Sheet1.Range("AP3")
                .Value = iScore
                .Font.Bold = True
            End With
            If iScore Mod 12 = 0 And iScore <> 0 Then
                Call GameBoard(2)
            End If
        End If
    Loop While iApples < iTotalApples
    
End Sub
