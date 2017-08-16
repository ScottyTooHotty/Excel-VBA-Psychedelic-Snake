VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)

    With Application
        .OnKey "{UP}"
        .OnKey "{DOWN}"
        .OnKey "{LEFT}"
        .OnKey "{RIGHT}"
        .OnKey "{BACKSPACE}"
        .OnKey "{ }"
    End With
    
End Sub

Private Sub Workbook_Open()

    With Application
        .OnKey "{BACKSPACE}", "EndTimer"
        .OnKey "{ }", "StartGame"
    End With

End Sub
