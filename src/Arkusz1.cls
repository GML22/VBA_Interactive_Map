VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Arkusz1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Sub kooo()
'
'  Dim i As Integer
'
'
'  For i = 2 To 400
'
'    If Cells(i, 1) = "pucki" Or Cells(i, 1) = "Gdynia" Or Cells(i, 1) = "Sopot" Or Cells(i, 1) = "Gdañsk" Or Cells(i, 1) = "S³upsk" Or Cells(i, 1) = "wejherowski" Or Cells(i, 1) = "lêborski" Or Cells(i, 1) = "s³upski" _
'             Or Cells(i, 1) = "bytowski" Or Cells(i, 1) = "cz³uchowski" Or Cells(i, 1) = "chojnicki" Or Cells(i, 1) = "koœcierski" Or Cells(i, 1) = "kartuski" Or Cells(i, 1) = "gdañski" Or Cells(i, 1) = "starogardzki" _
'              Or Cells(i, 1) = "tczewski" Or Cells(i, 1) = "nowodworski (pomorskie)" Or Cells(i, 1) = "malborski" Or Cells(i, 1) = "sztumski" Or Cells(i, 1) = "kwidzyñski" Then
'
'              Cells(i, 2) = "pomorskie"
'
'   ElseIf Cells(i, 1) = "s³awieñski" Or Cells(i, 1) = "Szczecin" Or Cells(i, 1) = "Koszalin" Or Cells(i, 1) = "bia³ogardzki" Or Cells(i, 1) = "koszaliñski" Or Cells(i, 1) = _
'            "szczecinecki" Or Cells(i, 1) = "wa³ecki" Or Cells(i, 1) = "drawski" Or Cells(i, 1) = "œwidwiñski" Or Cells(i, 1) = "ko³obrzeski" Or Cells(i, 1) = "gryficki" Or Cells(i, 1) = _
'            "³obeski" Or Cells(i, 1) = "Œwinoujœcie" Or Cells(i, 1) = "goleniowski" Or Cells(i, 1) = "stargardzki" Or Cells(i, 1) = "choszczeñski" Or Cells(i, 1) = _
'            "pyrzycki" Or Cells(i, 1) = "myœliborski" Or Cells(i, 1) = "gryfiñski" Or Cells(i, 1) = "policki" Or Cells(i, 1) = "kamieñski" Then
'
'        Cells(i, 2) = "zachodniopomorskie"
'
'   ElseIf Cells(i, 1) = "Elbl¹g" Or Cells(i, 1) = "Olsztyn" Or Cells(i, 1) = "braniewski" Or Cells(i, 1) = "elbl¹ski" Or Cells(i, 1) = "i³awski" Or Cells(i, 1) = _
'            "nowomiejski" Or Cells(i, 1) = "dzia³dowski" Or Cells(i, 1) = "ostródzki" Or Cells(i, 1) = "lidzbarski" Or Cells(i, 1) = "bartoszycki" Or Cells(i, 1) = _
'            "olsztyñski" Or Cells(i, 1) = "nidzicki" Or Cells(i, 1) = "szczycieñski" Or Cells(i, 1) = "mr¹gowski" Or Cells(i, 1) = "kêtrzyñski" Or Cells(i, 1) = _
'            "wêgorzewski" Or Cells(i, 1) = "go³dapski" Or Cells(i, 1) = "olecki" Or Cells(i, 1) = "gi¿ycki" Or Cells(i, 1) = "e³cki" Or Cells(i, 1) = "piski" Then
'            Cells(i, 2) = "warmiñsko-mazurskie"
'
'    ElseIf Cells(i, 1) = "sêpoleñski" Or Cells(i, 1) = "tucholski" Or Cells(i, 1) = "Bydgoszcz" Or Cells(i, 1) = "Toruñ" Or Cells(i, 1) = "Grudzi¹dz" Or Cells(i, 1) = "W³oc³awek" Or Cells(i, 1) = "che³miñski" Or Cells(i, 1) = _
'            "œwiecki" Or Cells(i, 1) = "grudzi¹dzki" Or Cells(i, 1) = "w¹brzeski" Or Cells(i, 1) = "brodnicki" Or Cells(i, 1) = "rypiñski" Or Cells(i, 1) = _
'            "golubsko-dobrzyñski" Or Cells(i, 1) = "lipnowski" Or Cells(i, 1) = "w³oc³awski" Or Cells(i, 1) = "radziejowski" Or Cells(i, 1) = _
'            "aleksandrowski" Or Cells(i, 1) = "toruñski" Or Cells(i, 1) = "inowroc³awski" Or Cells(i, 1) = "mogileñski" Or Cells(i, 1) = "bydgoski" Or Cells(i, 1) = _
'            "¿niñski" Then Cells(i, 2) = "kujawsko-pomorskie"
'
'    ElseIf Cells(i, 1) = "Poznañ" Or Cells(i, 1) = "Konin" Or Cells(i, 1) = "Kalisz" Or Cells(i, 1) = "Leszno" Or Cells(i, 1) = "z³otowski" Or Cells(i, 1) = "pilski" Or Cells(i, 1) = "nakielski" Or Cells(i, 1) = _
'            "chodzieski" Or Cells(i, 1) = "w¹growiecki" Or Cells(i, 1) = "obornicki" Or Cells(i, 1) = "czarnkowsko-trzcianecki" Or Cells(i, 1) = _
'            "szamotulski" Or Cells(i, 1) = "miêdzychodzki" Or Cells(i, 1) = "nowotomyski" Or Cells(i, 1) = "wolsztyñski" Or Cells(i, 1) = "grodziski (wielkopolskie)" Or Cells(i, 1) = _
'            "poznañski" Or Cells(i, 1) = "gnieŸnieñski" Or Cells(i, 1) = "kolski" Or Cells(i, 1) = "koniñski" Or Cells(i, 1) = "s³upecki" Or Cells(i, 1) = "wrzesiñski" Or Cells(i, 1) = _
'            "œremski" Or Cells(i, 1) = "jarociñski" Or Cells(i, 1) = "œredzki (wielkopolskie)" Or Cells(i, 1) = "leszczyñski" Or Cells(i, 1) = "gostyñski" _
'             Or Cells(i, 1) = "rawicki" Or Cells(i, 1) = "pleszewski" Or Cells(i, 1) = "kaliski" Or Cells(i, 1) = "ostrowski (wielkopolskie)" Or Cells(i, 1) = _
'            "krotoszyñski" Or Cells(i, 1) = "ostrzeszowski" Or Cells(i, 1) = "kêpiñski" Or Cells(i, 1) = "turecki" Or Cells(i, 1) = "koœciañski" Then
'
'            Cells(i, 2) = "wielkopolskie"
'
'    ElseIf Cells(i, 1) = "strzelecko-drezdenecki" Or Cells(i, 1) = "Gorzów Wielkopolski" Or Cells(i, 1) = "Zielona Góra" Or Cells(i, 1) = "gorzowski" Or Cells(i, 1) = _
'            "sulêciñski" Or Cells(i, 1) = "s³ubicki" Or Cells(i, 1) = "miêdzyrzecki" Or Cells(i, 1) = "œwiebodziñski" Or Cells(i, 1) = "zielonogórski" Or Cells(i, 1) = _
'            "kroœnieñski (lubuskie)" Or Cells(i, 1) = "¿arski" Or Cells(i, 1) = "nowosolski" Or Cells(i, 1) = "¿agañski" Or Cells(i, 1) = "wschowski" Then Cells(i, 2) = "lubuskie"
'
'    ElseIf Cells(i, 1) = "Wroc³aw" Or Cells(i, 1) = "Jelenia Góra" Or Cells(i, 1) = "Legnica" Or Cells(i, 1) = "wa³brzyski" Or Cells(i, 1) = "kamiennogórski" Or Cells(i, 1) = "z³otoryjski" _
'             Or Cells(i, 1) = "lubañski" Or Cells(i, 1) = "zgorzelecki" Or Cells(i, 1) = "boles³awiecki" Or Cells(i, 1) = "polkowicki" Or Cells(i, 1) = "g³ogowski" Or Cells(i, 1) = _
'            "górowski" Or Cells(i, 1) = "wo³owski" Or Cells(i, 1) = "legnicki" Or Cells(i, 1) = "jaworski" Or Cells(i, 1) = "trzebnicki" Or Cells(i, 1) = "milicki" Or Cells(i, 1) = _
'            "oleœnicki" Or Cells(i, 1) = "wroc³awski" Or Cells(i, 1) = "o³awski" Or Cells(i, 1) = "strzeliñski" Or Cells(i, 1) = "dzier¿oniowski" Or Cells(i, 1) = _
'            "z¹bkowicki" Or Cells(i, 1) = "lubiñski" Or Cells(i, 1) = "œredzki (dolnyœl¹sk)" Or Cells(i, 1) = "œwidnicki (dolnoœl¹skie)" Or Cells(i, 1) = _
'            "lwówecki" Or Cells(i, 1) = "jeleniogórski" Or Cells(i, 1) = "k³odzki" Then Cells(i, 2) = "dolnoœl¹skie"
'
'    ElseIf Cells(i, 1) = "Suwa³ki" Or Cells(i, 1) = "Bia³ystok" Or Cells(i, 1) = "£om¿a" Or Cells(i, 1) = "suwalski" Or Cells(i, 1) = "sejneñski" Or Cells(i, 1) = "augustowski" Or Cells(i, 1) = _
'            "grajewski" Or Cells(i, 1) = "kolneñski" Or Cells(i, 1) = "moniecki" Or Cells(i, 1) = "sokólski" Or Cells(i, 1) = "bia³ostocki" Or Cells(i, 1) = "³om¿yñski" Or Cells(i, 1) = _
'            "zambrowski" Or Cells(i, 1) = "wysokomazowiecki" Or Cells(i, 1) = "bielski (podlaskie)" Or Cells(i, 1) = "hajnowski" Or Cells(i, 1) = _
'            "siemiatycki" Then Cells(i, 2) = "podlaskie"
'
'    ElseIf Cells(i, 1) = "Ostro³êka" Or Cells(i, 1) = "P³ock" Or Cells(i, 1) = "Warszawa" Or Cells(i, 1) = "Radom" Or Cells(i, 1) = "Siedlce" Or Cells(i, 1) = "ostro³êcki" Or Cells(i, 1) = "przasnyski" Or Cells(i, 1) = "ciechanowski" Or Cells(i, 1) = _
'            "m³awski" Or Cells(i, 1) = "¿uromiñski" Or Cells(i, 1) = "sierpecki" Or Cells(i, 1) = "gostyniñski" Or Cells(i, 1) = "p³ocki" Or Cells(i, 1) = "p³oñski" Or Cells(i, 1) = _
'            "pu³tuski" Or Cells(i, 1) = "nowodworski (mazowieckie)" Or Cells(i, 1) = "sochaczewski" Or Cells(i, 1) = "wyszkowski" Or Cells(i, 1) = _
'            "ostrowski (mazowieckie)" Or Cells(i, 1) = "soko³owski" Or Cells(i, 1) = "³osicki" Or Cells(i, 1) = "siedlecki" Or Cells(i, 1) = "wêgrowski" Or Cells(i, 1) = _
'            "wo³omiñski" Or Cells(i, 1) = "miñski" Or Cells(i, 1) = "warszawski zachodni" Or Cells(i, 1) = "pruszkowski" Or Cells(i, 1) = "grodziski (mazowieckie)" Or Cells(i, 1) = _
'            "¿yrardowski" Or Cells(i, 1) = "piaseczyñski" Or Cells(i, 1) = "otwocki" Or Cells(i, 1) = "grójecki" Or Cells(i, 1) = "przysuski" Or Cells(i, 1) = _
'            "bia³obrzeski" Or Cells(i, 1) = "kozienicki" Or Cells(i, 1) = "garwoliñski" Or Cells(i, 1) = "zwoleñski" Or Cells(i, 1) = "szyd³owiecki" Or Cells(i, 1) = _
'            "makowski" Or Cells(i, 1) = "legionowski" Or Cells(i, 1) = "radomski" Or Cells(i, 1) = "lipski" Then Cells(i, 2) = "mazowieckie"
'
'    ElseIf Cells(i, 1) = "rawski" Or Cells(i, 1) = "£ódŸ" Or Cells(i, 1) = "Piotrków Trybunalski" Or Cells(i, 1) = "Skierniewice" Or Cells(i, 1) = "skierniewicki" Or Cells(i, 1) = "³owicki" Or Cells(i, 1) = _
'            "kutnowski" Or Cells(i, 1) = "³êczycki" Or Cells(i, 1) = "zgierski" Or Cells(i, 1) = "brzeziñski" Or Cells(i, 1) = "tomaszowski (³ódzkie)" Or Cells(i, 1) = _
'            "opoczyñski" Or Cells(i, 1) = "³ódzki wschodni" Or Cells(i, 1) = "piotrkowski" Or Cells(i, 1) = "radomszczañski" Or Cells(i, 1) = "pajêczañski" _
'             Or Cells(i, 1) = "be³chatowski" Or Cells(i, 1) = "pabianicki" Or Cells(i, 1) = "³aski" Or Cells(i, 1) = "zduñskowolski" Or Cells(i, 1) = "poddêbicki" Or Cells(i, 1) = _
'            "sieradzki" Or Cells(i, 1) = "wieluñski" Or Cells(i, 1) = "wieruszowski" Then Cells(i, 2) = "³ódzkie"
'
'    ElseIf Cells(i, 1) = "Opole" Or Cells(i, 1) = "g³ubczycki" Or Cells(i, 1) = "kêdzierzyñsko-kozielski" Or Cells(i, 1) = _
'            "strzelecki" Or Cells(i, 1) = "oleski" Or Cells(i, 1) = "kluczborski" Or Cells(i, 1) = "krapkowicki" Or Cells(i, 1) = "prudnicki" Or Cells(i, 1) = _
'            "opolski (opolskie)" Or Cells(i, 1) = "namys³owski" Or Cells(i, 1) = "brzeski (opolskie)" Or Cells(i, 1) = "nyski" Then Cells(i, 2) = "opolskie"
'
'    ElseIf Cells(i, 1) = "Czêstochowa" Or Cells(i, 1) = "Bielsko-Bia³a" Or Cells(i, 1) = "Ruda Œl¹ska" Or Cells(i, 1) = "Œwiêtoch³owice" Or Cells(i, 1) = "Chorzów" _
'         Or Cells(i, 1) = "Siemianowice Œl¹skie" Or Cells(i, 1) = "Piekary Œl¹skie" Or Cells(i, 1) = "Bytom" Or Cells(i, 1) = "Zabrze" Or Cells(i, 1) = "Gliwice" Or Cells(i, 1) = "Rybnik" Or Cells(i, 1) = "Jastrzêbie-Zdrój" Or Cells(i, 1) = "¯ory" Or Cells(i, 1) = _
'        "Tychy" Or Cells(i, 1) = "Jaworzno" Or Cells(i, 1) = "Mys³owice" Or Cells(i, 1) = "Katowice" Or Cells(i, 1) = "Sosnowiec" Or Cells(i, 1) = "D¹browa-Górnicza" Or Cells(i, 1) = "bieruñsko-lêdziñski" Or Cells(i, 1) = "pszczyñski" Or Cells(i, 1) = _
'        "bielski (œl¹skie)" Or Cells(i, 1) = "cieszyñski" Or Cells(i, 1) = "wodzis³awski" Or Cells(i, 1) = "rybnicki" Or Cells(i, 1) = "raciborski" Or Cells(i, 1) = _
'        "gliwicki" Or Cells(i, 1) = "miko³owski" Or Cells(i, 1) = "bêdziñski" Or Cells(i, 1) = "tarnogórski" Or Cells(i, 1) = "zawierciañski" Or Cells(i, 1) = _
'        "myszkowski" Or Cells(i, 1) = "czêstochowski" Or Cells(i, 1) = "lubliniecki" Or Cells(i, 1) = "k³obucki" Or Cells(i, 1) = "¿ywiecki" Then Cells(i, 2) = "œl¹skie"
'
'    ElseIf Cells(i, 1) = "Kielce" Or Cells(i, 1) = "konecki" Or Cells(i, 1) = "skar¿yski" Or Cells(i, 1) = "starachowicki" Or Cells(i, 1) = _
'            "ostrowiecki" Or Cells(i, 1) = "opatowski" Or Cells(i, 1) = "sandomierski" Or Cells(i, 1) = "staszowski" Or Cells(i, 1) = "kielecki" Or Cells(i, 1) = _
'            "w³oszczowski" Or Cells(i, 1) = "jêdrzejowski" Or Cells(i, 1) = "piñczowski" Or Cells(i, 1) = "buski" Or Cells(i, 1) = "kazimierski" Then Cells(i, 2) = "œwiêtokrzyskie"
'
'    ElseIf Cells(i, 1) = "Zamoœæ" Or Cells(i, 1) = "Che³m" Or Cells(i, 1) = "Bia³a Podlaska" Or Cells(i, 1) = "Lublin" Or Cells(i, 1) = "bialski" Or Cells(i, 1) = "radzyñski" Or Cells(i, 1) = "³ukowski" Or Cells(i, 1) = "rycki" Or Cells(i, 1) = _
'            "parczewski" Or Cells(i, 1) = "w³odawski" Or Cells(i, 1) = "lubartowski" Or Cells(i, 1) = "pu³awski" Or Cells(i, 1) = "opolski (lubelskie)" Or Cells(i, 1) = _
'            "lubelski" Or Cells(i, 1) = "kraœnicki" Or Cells(i, 1) = "œwidnicki (lubelskie)" Or Cells(i, 1) = "che³mski" Or Cells(i, 1) = "krasnostawski" Or Cells(i, 1) = _
'            "hrubieszowski" Or Cells(i, 1) = "zamojski" Or Cells(i, 1) = "janowski" Or Cells(i, 1) = "bi³gorajski" Or Cells(i, 1) = "³êczyñski" Or Cells(i, 1) = _
'            "tomaszowski (lubelskie)" Then Cells(i, 2) = "lubelskie"
'
'    ElseIf Cells(i, 1) = "Rzeszów" Or Cells(i, 1) = "Przemyœl" Or Cells(i, 1) = "Krosno" Or Cells(i, 1) = "Tarnobrzeg" Or Cells(i, 1) = "jaros³awski" Or Cells(i, 1) = "przemyski" Or Cells(i, 1) = "le¿ajski" Or Cells(i, 1) = _
'            "przeworski" Or Cells(i, 1) = "³añcucki" Or Cells(i, 1) = "ni¿añski" Or Cells(i, 1) = "stalowowolski" Or Cells(i, 1) = "tarnobrzeski" Or Cells(i, 1) = _
'            "kolbuszowski" Or Cells(i, 1) = "rzeszowski" Or Cells(i, 1) = "ropczycko-sêdziszowski" Or Cells(i, 1) = "mielecki" Or Cells(i, 1) = "dêbicki" Or Cells(i, 1) = _
'            "brzozowski" Or Cells(i, 1) = "kroœnieñski" Or Cells(i, 1) = "jasielski" Or Cells(i, 1) = "sanocki" Or Cells(i, 1) = "leski" Or Cells(i, 1) = "bieszczadzki" Or Cells(i, 1) = _
'            "strzy¿owski" Or Cells(i, 1) = "lubaczowski" Then Cells(i, 2) = "podkarpackie"
'
'   ElseIf Cells(i, 1) = "Kraków" Or Cells(i, 1) = "Tarnów" Or Cells(i, 1) = "Nowy S¹cz" Or Cells(i, 1) = "d¹browski" Or Cells(i, 1) = "gorlicki" Or Cells(i, 1) = "nowos¹decki" Or Cells(i, 1) = _
'            "bocheñski" Or Cells(i, 1) = "brzeski (ma³opolskie)" Or Cells(i, 1) = "tarnowski" Or Cells(i, 1) = "limanowski" Or Cells(i, 1) = "wielicki" Or Cells(i, 1) = _
'            "myœlenicki" Or Cells(i, 1) = "suski" Or Cells(i, 1) = "proszowicki" Or Cells(i, 1) = "miechowski" Or Cells(i, 1) = "krakowski" Or Cells(i, 1) = "wadowicki" Or Cells(i, 1) = _
'            "oœwiêcimski" Or Cells(i, 1) = "chrzanowski" Or Cells(i, 1) = "okulski" Or Cells(i, 1) = "nowotarski" Or Cells(i, 1) = "tatrzañski" Then Cells(i, 2) = "ma³opolskie"
'
'   End If
'
'  Next i
'
'End Sub

'Sub ludnosc()
'
'   Dim path, nazwa, nazwa11 As String
'   Dim i, j, liczba As Integer
'
'   For i = 178 To 178
'
'        nazwa = Sheets("Excele").Cells(i, 1)
'        'If InStr(1, Mid(nazwa, 6, Len(nazwa) - 9), "_", vbTextCompare) <> 0 Then nazwa1 = Replace(Mid(nazwa, 6, Len(nazwa) - 9), "_", " ")
'        Set findrange = Sheets("Mapka Dane").Range("A:A")
'        Set foundcell1 = findrange.Find(What:=Mid(nazwa, 6, Len(nazwa) - 9), LookIn:=xlValues, LookAt:=xlWhole)
'        If Not foundcell1 Is Nothing Then
'        If Cells(foundcell1.row, 3) = "" Then
'
'            path = "C:\Users\Mateusz\Desktop\Mapa Polski interaktywna\powiaty ludnoœæ\" & Sheets("Excele").Cells(i, 1)
'
'            Workbooks.Open Filename:=path
'
'            Sheets("Roczniki ludnoœci").Select
'
'            For j = 1 To 25
'
'                Windows(nazwa).Activate
'
'                liczba = 7 + 87 * (j - 1)
'
'                Sheets("Roczniki ludnoœci").Range("D" & liczba & ":" & "F" & liczba).Select
'                Selection.Copy
'
'                Windows("Mapka interaktywna 3.3.xlsm").Activate
'                Sheets("Mapka dane").Cells(foundcell1.row, 3 + 3 * (j - 1)).Select
'                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
'                :=False, Transpose:=False
'
'            Next j
'
'            Windows(nazwa).Activate
'            ActiveWorkbook.Close
'            Windows("Mapka interaktywna 3.3.xlsm").Activate
'
'        End If
'        End If
'
'    Next i
'
'End Sub





