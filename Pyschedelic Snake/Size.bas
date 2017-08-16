Attribute VB_Name = "Size"
Option Explicit

Public bStillGrowing As Boolean
Public iStillGrowing As Integer

Sub Eat(bEat As Boolean)

    If bEat = True Then
        iStillGrowing = iStillGrowing + 2
    End If

    lLength = lLength + 1
    ReDim Preserve rBody(lStart To lLength)
    
    If iStillGrowing = 0 Then
        bStillGrowing = False
    Else
        bStillGrowing = True
        iStillGrowing = iStillGrowing - 1
    End If
    
End Sub
