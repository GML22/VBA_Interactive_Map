VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ten_skoroszyt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)

    Application.ScreenUpdating = True

End Sub

Private Sub Workbook_Open()

    Odswiez
    'RefreshCharts  'odswieza wykresy, ¿eby update'owa³y siê w czasie dzia³ania excela
    Application.ScreenUpdating = False
    'Application.Calculation = xlCalculationManual
    ZoomToRange Range("Plansza"), False
    Range("A1").Select
   
End Sub

Sub fff()
ActiveWindow.DisplayWorkbookTabs = False
ActiveWindow.DisplayHeadings = False


End Sub


