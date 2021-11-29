VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub ccdgf()

 Set obj = CreateObject("InternetExplorer.Application")
    
    obj.Navigate ("http://pl.wikipedia.org/wiki/Powiat_grodziski_(województwo_wielkopolskie)")   ' otwiera i wchodzi na strone imdb
    
    Do Until obj.ReadyState = 4 'czeka a¿ strona sie za³aduje, 4 oznacza za³adowan¹
    Loop
       
    Set ElementCol = obj.Document.getElementsByTagName("a") 'uzyskuje linki z adresu
          
    i = 1
    
    For Each Link In ElementCol
           
          Sheets("Arkusz1").Cells(i, 1).Value = Link.href
          
'          Set Znajdz = Sheets("Testowy").Range("A" & i).Find(What:="http://www.imdb.com/title/", LookIn:=xlValues, LookAt _
'          :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False, SearchFormat:=False)
'
'        If Not Znajdz Is Nothing Then
'
'           adres = Sheets("Testowy").Cells(Znajdz.row, 1)
'           Sheets("Testowy").Cells.Clear      'czyœci ca³y arkusz
'            Exit For
'
'        End If
'
          i = i + 1
          
    Next Link

MsgBox "dd"

End Sub

Sub hhhh()


Sheets("Arkusz1").Shapes.AddPicture "http://upload.wikimedia.org/wikipedia/commons/thumb/3/30/POL_wojew%C3%B3dztwo_wielkopolskie_powiat_grodziski_map.svg/476px-POL_wojew%C3%B3dztwo_wielkopolskie_powiat_grodziski_map.svg.png" _
                          , msoFalse, msoTrue, 100, 100, 40, 60

End Sub

Sub dddd()

ActiveChart.Parent.Name = "Wykres 784"

End Sub

Sub ffgg()

MapkaNum

End Sub
