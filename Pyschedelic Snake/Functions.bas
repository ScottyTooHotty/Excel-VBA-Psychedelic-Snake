Attribute VB_Name = "Functions"
Option Explicit

Public Function CellHasBorder(rCell As Range)

    CellHasBorder = False
    
    If rCell.Borders(xlEdgeTop).LineStyle <> xlNone And _
        rCell.Borders(xlEdgeBottom).LineStyle <> xlNone And _
        rCell.Borders(xlEdgeLeft).LineStyle <> xlNone And _
        rCell.Borders(xlEdgeRight).LineStyle <> xlNone Then
        CellHasBorder = True
    End If
    
End Function

Public Function RandomColour() As Integer

    Randomize
    
    RandomColour = Int(56 * Rnd) + 1

End Function
