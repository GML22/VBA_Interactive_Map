VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub jjj()

    Dim i, j As Integer

    For i = 1 To 57

        For j = 1 To 70

            Sheets("Mapy pomocnicze").Cells(i + 180, j) = WojNum(i, j, 0)

        Next j

    Next i

End Sub






