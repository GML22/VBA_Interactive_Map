Attribute VB_Name = "Suby"

Public Sub CalyKraj()

    If ActiveSheet.Shapes.Range(Array("Kraj")).Glow.Radius <> 0 Then Exit Sub
        
    Dim i As Integer
    
    If ActiveSheet.Shapes.Range(Array("Powiat")).Glow.Radius <> 0 Then
    
        With ActiveSheet.Shapes("pole tekstowe 2")
    
            .Visible = False
            .TextFrame.Characters.Text = ""
    
        End With
        
        counter = ""
        
        With Sheets("Powiaty").Shapes(counter)

            .Fill.Transparency = 0

        End With

        ActiveSheet.Shapes.Range(Array("Powiat")).Glow.Radius = 0
        ActiveSheet.Shapes.Range(Array("Woj")).Glow.Radius = 0

        With ActiveSheet.Shapes.Range(Array("Kraj")).Glow
            .color.ObjectThemeColor = msoThemeColorBackground1
            ' .Transparency = 0.2
            .Radius = 5
        End With
        
        Range("H42:K43").FormulaR1C1 = ""
        Range("A44:S55").Borders(xlEdgeLeft).LineStyle = xlNone
        Range("A44:S55").Borders(xlEdgeTop).LineStyle = xlNone
        Range("A44:S55").Borders(xlEdgeRight).LineStyle = xlNone
        
        ActiveSheet.Shapes.Range(Array("Ruda Œl¹ska ", "Œwiêtoch³owice ", "Chorzów " _
        , "bêdziñski ", "Siemianowice Œl¹skie ", "Piekary Œl¹skie ", "Bytom ", _
        "Zabrze ", "Gliwice ", "Rybnik ", "Jastrzêbie-Zdrój ", "miko³owski ", "¯ory ", _
        "Tychy ", "bieruñsko-lêdziñski ", "Jaworzno ", "Mys³owice ", "Katowice ", _
        "Sosnowiec ", "D¹browa-Górnicza ")).Group.Name = " "
        
        On Error Resume Next
    
        ActiveSheet.Shapes.Range(Array(" ")).Visible = msoFalse
          
         With ActiveSheet.Shapes.Range(Array( _
        "hajnowski", "Toruñ", "miêdzyrzecki", "sulêciñski", "pucki", "wejherowski", "lêborski", "s³upski", "s³awieñski", "Gdynia", "S³upsk", "kartuski", "suwalski", "bytowski", "Gdañsk", "braniewski", "go³dapski", "bartoszycki", "wêgorzewski", "sejneñski", "kêtrzyñski", "koszaliñski", "nowodworski (pomorskie)", _
        "gdañski", "elbl¹ski", "Elbl¹g", "olecki", "ko³obrzeski", "Koszalin", "lidzbarski", "Suwa³ki", "gryficki", "gi¿ycki", "koœcierski", "malborski", "bia³ogardzki", "tczewski", "starogardzki", "kamieñski", "olsztyñski", "augustowski", "Œwinoujœcie", "ostródzki", "e³cki", "chojnicki", "sztumski", "szczecinecki", "cz³uchowski", "œwidwiñski", "mr¹gowski", _
        "piski", "³obeski", "i³awski", "kwidzyñski", "goleniowski", "Olsztyn", "grajewski", "policki", "szczycieñski", "sokólski", "tucholski", "moniecki", "drawski", "œwiecki", "z³otowski", "sêpoleñski", "grudzi¹dzki", "Szczecin", "stargardzki", "kolneñski", "nowomiejski", "Grudzi¹dz", "nidzicki", "ostro³êcki", "³om¿yñski", "wa³ecki", "bia³ostocki", "brodnicki", "bydgoski", "dzia³dowski", "gryfiñski", "che³miñski", "w¹brzeski", "choszczeñski", _
        "przasnyski", "pilski", "pyrzycki", "nakielski", "m³awski", "toruñski", "Bia³ystok", "golubsko-dobrzyñski", "zambrowski", "£om¿a", "rypiñski", "Bydgoszcz", "¿uromiñski", "wysokomazowiecki", "strzelecko-drezdenecki", "czarnkowsko-trzcianecki", "makowski", "Ostro³êka", "ciechanowski", "myœliborski", "chodzieski")) _
        .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    
        With ActiveSheet.Shapes.Range(Array( _
        "hajnowski", "Toruñ", "miêdzyrzecki", "sulêciñski", "pucki", "wejherowski", "lêborski", "s³upski", "s³awieñski", "Gdynia", "S³upsk", "kartuski", "suwalski", "bytowski", "Gdañsk", "braniewski", "go³dapski", "bartoszycki", "wêgorzewski", "sejneñski", "kêtrzyñski", "koszaliñski", "nowodworski (pomorskie)", _
        "gdañski", "elbl¹ski", "Elbl¹g", "olecki", "ko³obrzeski", "Koszalin", "lidzbarski", "Suwa³ki", "gryficki", "gi¿ycki", "koœcierski", "malborski", "bia³ogardzki", "tczewski", "starogardzki", "kamieñski", "olsztyñski", "augustowski", "Œwinoujœcie", "ostródzki", "e³cki", "chojnicki", "sztumski", "szczecinecki", "cz³uchowski", "œwidwiñski", "mr¹gowski", _
        "piski", "³obeski", "i³awski", "kwidzyñski", "goleniowski", "Olsztyn", "grajewski", "policki", "szczycieñski", "sokólski", "tucholski", "moniecki", "drawski", "œwiecki", "z³otowski", "sêpoleñski", "grudzi¹dzki", "Szczecin", "stargardzki", "kolneñski", "nowomiejski", "Grudzi¹dz", "nidzicki", "ostro³êcki", "³om¿yñski", "wa³ecki", "bia³ostocki", "brodnicki", "bydgoski", "dzia³dowski", "gryfiñski", "che³miñski", "w¹brzeski", "choszczeñski", _
        "przasnyski", "pilski", "pyrzycki", "nakielski", "m³awski", "toruñski", "Bia³ystok", "golubsko-dobrzyñski", "zambrowski", "£om¿a", "rypiñski", "Bydgoszcz", "¿uromiñski", "wysokomazowiecki", "strzelecko-drezdenecki", "czarnkowsko-trzcianecki", "makowski", "Ostro³êka", "ciechanowski", "myœliborski", "chodzieski")) _
        .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
            .Visible = msoTrue
            .Weight = 2.5
        End With
        
        ActiveSheet.Shapes.Range(Array( _
        "hajnowski", "Toruñ", "miêdzyrzecki", "sulêciñski", "pucki", "wejherowski", "lêborski", "s³upski", "s³awieñski", "Gdynia", "S³upsk", "kartuski", "suwalski", "bytowski", "Gdañsk", "braniewski", "go³dapski", "bartoszycki", "wêgorzewski", "sejneñski", "kêtrzyñski", "koszaliñski", "nowodworski (pomorskie)", _
        "gdañski", "elbl¹ski", "Elbl¹g", "olecki", "ko³obrzeski", "Koszalin", "lidzbarski", "Suwa³ki", "gryficki", "gi¿ycki", "koœcierski", "malborski", "bia³ogardzki", "tczewski", "starogardzki", "kamieñski", "olsztyñski", "augustowski", "Œwinoujœcie", "ostródzki", "e³cki", "chojnicki", "sztumski", "szczecinecki", "cz³uchowski", "œwidwiñski", "mr¹gowski", _
        "piski", "³obeski", "i³awski", "kwidzyñski", "goleniowski", "Olsztyn", "grajewski", "policki", "szczycieñski", "sokólski", "tucholski", "moniecki", "drawski", "œwiecki", "z³otowski", "sêpoleñski", "grudzi¹dzki", "Szczecin", "stargardzki", "kolneñski", "nowomiejski", "Grudzi¹dz", "nidzicki", "ostro³êcki", "³om¿yñski", "wa³ecki", "bia³ostocki", "brodnicki", "bydgoski", "dzia³dowski", "gryfiñski", "che³miñski", "w¹brzeski", "choszczeñski", _
        "przasnyski", "pilski", "pyrzycki", "nakielski", "m³awski", "toruñski", "Bia³ystok", "golubsko-dobrzyñski", "zambrowski", "£om¿a", "rypiñski", "Bydgoszcz", "¿uromiñski", "wysokomazowiecki", "strzelecko-drezdenecki", "czarnkowsko-trzcianecki", "makowski", "Ostro³êka", "ciechanowski", "myœliborski", "chodzieski")) _
        .Group.Name = "Grupa 1"
        
        With ActiveSheet.Shapes.Range(Array("¿niñski", "w¹growiecki", "Konin", "Sopot", "Siedlce", "ostrowski (mazowieckie)", "lipnowski", "bielski (podlaskie)", "inowroc³awski", "sierpecki", "gorzowski", "aleksandrowski", "p³oñski", "obornicki", "wyszkowski", "miêdzychodzki", "Gorzów Wielkopolski", "mogileñski", "pu³tuski", "szamotulski", "p³ocki", "w³oc³awski", "siemiatycki", "gnieŸnieñski", "W³oc³awek", "soko³owski", "wêgrowski", "radziejowski", "poznañski", "nowodworski (mazowieckie)", "s³ubicki", "P³ock", "wo³omiñski", "legionowski", "s³upecki", "nowotomyski", "koniñski", "Poznañ", "gostyniñski", "siedlecki", "³osicki", "œwiebodziñski", "wrzesiñski", "kolski", "miñski", "warszawski zachodni", "kutnowski", "sochaczewski", "Warszawa")) _
         .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    
         With ActiveSheet.Shapes.Range(Array("¿niñski", "w¹growiecki", "Konin", "Sopot", "Siedlce", "ostrowski (mazowieckie)", "lipnowski", "bielski (podlaskie)", "inowroc³awski", "sierpecki", "gorzowski", "aleksandrowski", "p³oñski", "obornicki", "wyszkowski", "miêdzychodzki", "Gorzów Wielkopolski", "mogileñski", "pu³tuski", "szamotulski", "p³ocki", "w³oc³awski", "siemiatycki", "gnieŸnieñski", "W³oc³awek", "soko³owski", "wêgrowski", "radziejowski", "poznañski", "nowodworski (mazowieckie)", "s³ubicki", "P³ock", "wo³omiñski", "legionowski", "s³upecki", "nowotomyski", "koniñski", "Poznañ", "gostyniñski", "siedlecki", "³osicki", "œwiebodziñski", "wrzesiñski", "kolski", "miñski", "warszawski zachodni", "kutnowski", "sochaczewski", "Warszawa")) _
         .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
            .Visible = msoTrue
            .Weight = 2.5
        End With
        
        ActiveSheet.Shapes.Range(Array("¿niñski", "w¹growiecki", "Konin", "Sopot", "Siedlce", "ostrowski (mazowieckie)", "lipnowski", "bielski (podlaskie)", "inowroc³awski", "sierpecki", "gorzowski", "aleksandrowski", "p³oñski", "obornicki", "wyszkowski", "miêdzychodzki", "Gorzów Wielkopolski", "mogileñski", "pu³tuski", "szamotulski", "p³ocki", "w³oc³awski", "siemiatycki", "gnieŸnieñski", "W³oc³awek", "soko³owski", "wêgrowski", "radziejowski", "poznañski", "nowodworski (mazowieckie)", "s³ubicki", "P³ock", "wo³omiñski", "legionowski", "s³upecki", "nowotomyski", "koniñski", "Poznañ", "gostyniñski", "siedlecki", "³osicki", "œwiebodziñski", "wrzesiñski", "kolski", "miñski", "warszawski zachodni", "kutnowski", "sochaczewski", "Warszawa")) _
         .Group.Name = "Grupa 2"
        
        With ActiveSheet.Shapes.Range(Array( _
        "bialski", "œredzki (wielkopolskie)", "grodziski (wielkopolskie)", "³owicki", "kroœnieñski (lubuskie)", "wolsztyñski", "zielonogórski", "otwocki", "koœciañski", "pruszkowski", "³êczycki", "grodziski (mazowieckie)", "œremski", "turecki", "¿yrardowski", "piaseczyñski", "skierniewicki", "jarociñski", "Bia³a Podlaska", "³ukowski", "pleszewski", "poddêbicki", "nowosolski", "garwoliñski", "kaliski", "Zielona Góra", "wschowski", "zgierski", "radzyñski", "leszczyñski", "Skierniewice", "gostyñski", "grójecki", "brzeziñski", "¿arski", "krotoszyñski", "rawski", "Leszno", "kozienicki", "w³odawski", "£ódŸ", "parczewski", "sieradzki", "rycki", "g³ogowski", "¿agañski", "ostrowski (wielkopolskie)", "pabianicki", "górowski", "Kalisz", "tomaszowski (³ódzkie)", "zduñskowolski", "³aski", "bia³obrzeski", "lubartowski", "rawicki", "polkowicki", _
        "milicki", "piotrkowski", "przysuski", "lubiñski", "radomski", "pu³awski", "opoczyñski", "trzebnicki", "boles³awiecki", "ostrzeszowski", "be³chatowski", "wo³owski", "zgorzelecki", "Radom", "zwoleñski", "lubelski", "³êczyñski", "Piotrków Trybunalski", "che³mski", "oleœnicki", "wieruszowski", "legnicki", "wieluñski", "kêpiñski", "szyd³owiecki", "Lublin", "konecki", "opolski (lubelskie)", "œwidnicki (lubelskie)", "lipski", "z³otoryjski", "pajêczañski", "radomszczañski", "œredzki (dolnyœl¹sk)", "wroc³awski", "Legnica", "Che³m", "lwówecki", "lubañski", "starachowicki", "Wroc³aw", "skar¿yski", "namys³owski", "krasnostawski", "kluczborski", "jaworski", "oleski", "opatowski", "o³awski", "ostrowiecki", "kraœnicki", "k³obucki", "hrubieszowski", "kielecki", "jeleniogórski", "œwidnicki (dolnoœl¹skie)", "czêstochowski", "w³oszczowski", "brzeski (opolskie)", "zamojski", _
        "Jelenia Góra", "opolski (opolskie)", "janowski", "bi³gorajski", "Kielce", "kamiennogórski", "strzeliñski", "sandomierski", "wa³brzyski", "Czêstochowa", "jêdrzejowski", "lubliniecki", "stalowowolski", "dzier¿oniowski", "Zamoœæ", "tomaszowski (lubelskie)", "tarnobrzeski", "z¹bkowicki", "zawierciañski", "staszowski", "Opole", "strzelecki", "myszkowski", "k³odzki", "piñczowski", "Tarnobrzeg", "ni¿añski", "nyski", "buski", "tarnogórski", "krapkowicki", "prudnicki", "gliwicki", "bêdziñski", "mielecki", "miechowski", "kolbuszowski", "lubaczowski", "D¹browa-Górnicza", "okulski", "le¿ajski", "Bytom", "Piekary Œl¹skie", "kêdzierzyñsko-kozielski", "rzeszowski", "kazimierski", "Gliwice", "Zabrze", "d¹browski", "przeworski", "Siemianowice Œl¹skie", "Chorzów", "Ruda Œl¹ska", "proszowicki", "Œwiêtoch³owice", "Sosnowiec", "g³ubczycki", "krakowski", "Katowice", "³añcucki", _
        "jaros³awski", "Mys³owice", "raciborski", "miko³owski", "tarnowski", "chrzanowski", "brzeski (ma³opolskie)", "ropczycko-sêdziszowski", "rybnicki", "dêbicki", "Rybnik", "Tychy", "bieruñsko-lêdziñski", "bocheñski", "Kraków", "Rzeszów", "oœwiêcimski", "wielicki", "pszczyñski", "Tarnów", "wodzis³awski", "wadowicki", "przemyski", "strzy¿owski", "bielski (œl¹skie)", "cieszyñski", "myœlenicki", "jasielski", "Bielsko-Bia³a", "kroœnieñski", "brzozowski", "Przemyœl", "limanowski", "¿ywiecki", "suski", "nowos¹decki", "gorlicki", "Krosno", "bieszczadzki", "sanocki", "Nowy S¹cz", "nowotarski", "leski", "tatrzañski", "¯ory", "Jastrzêbie-Zdrój", "Jaworzno", "³ódzki wschodni")) _
        .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(220, 20, 60)
            .Solid
        End With
    
       With ActiveSheet.Shapes.Range(Array( _
        "bialski", "œredzki (wielkopolskie)", "grodziski (wielkopolskie)", "³owicki", "kroœnieñski (lubuskie)", "wolsztyñski", "zielonogórski", "otwocki", "koœciañski", "pruszkowski", "³êczycki", "grodziski (mazowieckie)", "œremski", "turecki", "¿yrardowski", "piaseczyñski", "skierniewicki", "jarociñski", "Bia³a Podlaska", "³ukowski", "pleszewski", "poddêbicki", "nowosolski", "garwoliñski", "kaliski", "Zielona Góra", "wschowski", "zgierski", "radzyñski", "leszczyñski", "Skierniewice", "gostyñski", "grójecki", "brzeziñski", "¿arski", "krotoszyñski", "rawski", "Leszno", "kozienicki", "w³odawski", "£ódŸ", "parczewski", "sieradzki", "rycki", "g³ogowski", "¿agañski", "ostrowski (wielkopolskie)", "pabianicki", "górowski", "Kalisz", "tomaszowski (³ódzkie)", "zduñskowolski", "³aski", "bia³obrzeski", "lubartowski", "rawicki", "polkowicki", _
        "milicki", "piotrkowski", "przysuski", "lubiñski", "radomski", "pu³awski", "opoczyñski", "trzebnicki", "boles³awiecki", "ostrzeszowski", "be³chatowski", "wo³owski", "zgorzelecki", "Radom", "zwoleñski", "lubelski", "³êczyñski", "Piotrków Trybunalski", "che³mski", "oleœnicki", "wieruszowski", "legnicki", "wieluñski", "kêpiñski", "szyd³owiecki", "Lublin", "konecki", "opolski (lubelskie)", "œwidnicki (lubelskie)", "lipski", "z³otoryjski", "pajêczañski", "radomszczañski", "œredzki (dolnyœl¹sk)", "wroc³awski", "Legnica", "Che³m", "lwówecki", "lubañski", "starachowicki", "Wroc³aw", "skar¿yski", "namys³owski", "krasnostawski", "kluczborski", "jaworski", "oleski", "opatowski", "o³awski", "ostrowiecki", "kraœnicki", "k³obucki", "hrubieszowski", "kielecki", "jeleniogórski", "œwidnicki (dolnoœl¹skie)", "czêstochowski", "w³oszczowski", "brzeski (opolskie)", "zamojski", _
        "Jelenia Góra", "opolski (opolskie)", "janowski", "bi³gorajski", "Kielce", "kamiennogórski", "strzeliñski", "sandomierski", "wa³brzyski", "Czêstochowa", "jêdrzejowski", "lubliniecki", "stalowowolski", "dzier¿oniowski", "Zamoœæ", "tomaszowski (lubelskie)", "tarnobrzeski", "z¹bkowicki", "zawierciañski", "staszowski", "Opole", "strzelecki", "myszkowski", "k³odzki", "piñczowski", "Tarnobrzeg", "ni¿añski", "nyski", "buski", "tarnogórski", "krapkowicki", "prudnicki", "gliwicki", "bêdziñski", "mielecki", "miechowski", "kolbuszowski", "lubaczowski", "D¹browa-Górnicza", "okulski", "le¿ajski", "Bytom", "Piekary Œl¹skie", "kêdzierzyñsko-kozielski", "rzeszowski", "kazimierski", "Gliwice", "Zabrze", "d¹browski", "przeworski", "Siemianowice Œl¹skie", "Chorzów", "Ruda Œl¹ska", "proszowicki", "Œwiêtoch³owice", "Sosnowiec", "g³ubczycki", "krakowski", "Katowice", "³añcucki", _
        "jaros³awski", "Mys³owice", "raciborski", "miko³owski", "tarnowski", "chrzanowski", "brzeski (ma³opolskie)", "ropczycko-sêdziszowski", "rybnicki", "dêbicki", "Rybnik", "Tychy", "bieruñsko-lêdziñski", "bocheñski", "Kraków", "Rzeszów", "oœwiêcimski", "wielicki", "pszczyñski", "Tarnów", "wodzis³awski", "wadowicki", "przemyski", "strzy¿owski", "bielski (œl¹skie)", "cieszyñski", "myœlenicki", "jasielski", "Bielsko-Bia³a", "kroœnieñski", "brzozowski", "Przemyœl", "limanowski", "¿ywiecki", "suski", "nowos¹decki", "gorlicki", "Krosno", "bieszczadzki", "sanocki", "Nowy S¹cz", "nowotarski", "leski", "tatrzañski", "¯ory", "Jastrzêbie-Zdrój", "Jaworzno", "³ódzki wschodni")) _
        .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(220, 20, 60)
            .Visible = msoTrue
            .Weight = 2.5
        End With
        
        ActiveSheet.Shapes.Range(Array( _
        "bialski", "œredzki (wielkopolskie)", "grodziski (wielkopolskie)", "³owicki", "kroœnieñski (lubuskie)", "wolsztyñski", "zielonogórski", "otwocki", "koœciañski", "pruszkowski", "³êczycki", "grodziski (mazowieckie)", "œremski", "turecki", "¿yrardowski", "piaseczyñski", "skierniewicki", "jarociñski", "Bia³a Podlaska", "³ukowski", "pleszewski", "poddêbicki", "nowosolski", "garwoliñski", "kaliski", "Zielona Góra", "wschowski", "zgierski", "radzyñski", "leszczyñski", "Skierniewice", "gostyñski", "grójecki", "brzeziñski", "¿arski", "krotoszyñski", "rawski", "Leszno", "kozienicki", "w³odawski", "£ódŸ", "parczewski", "sieradzki", "rycki", "g³ogowski", "¿agañski", "ostrowski (wielkopolskie)", "pabianicki", "górowski", "Kalisz", "tomaszowski (³ódzkie)", "zduñskowolski", "³aski", "bia³obrzeski", "lubartowski", "rawicki", "polkowicki", _
        "milicki", "piotrkowski", "przysuski", "lubiñski", "radomski", "pu³awski", "opoczyñski", "trzebnicki", "boles³awiecki", "ostrzeszowski", "be³chatowski", "wo³owski", "zgorzelecki", "Radom", "zwoleñski", "lubelski", "³êczyñski", "Piotrków Trybunalski", "che³mski", "oleœnicki", "wieruszowski", "legnicki", "wieluñski", "kêpiñski", "szyd³owiecki", "Lublin", "konecki", "opolski (lubelskie)", "œwidnicki (lubelskie)", "lipski", "z³otoryjski", "pajêczañski", "radomszczañski", "œredzki (dolnyœl¹sk)", "wroc³awski", "Legnica", "Che³m", "lwówecki", "lubañski", "starachowicki", "Wroc³aw", "skar¿yski", "namys³owski", "krasnostawski", "kluczborski", "jaworski", "oleski", "opatowski", "o³awski", "ostrowiecki", "kraœnicki", "k³obucki", "hrubieszowski", "kielecki", "jeleniogórski", "œwidnicki (dolnoœl¹skie)", "czêstochowski", "w³oszczowski", "brzeski (opolskie)", "zamojski", _
        "Jelenia Góra", "opolski (opolskie)", "janowski", "bi³gorajski", "Kielce", "kamiennogórski", "strzeliñski", "sandomierski", "wa³brzyski", "Czêstochowa", "jêdrzejowski", "lubliniecki", "stalowowolski", "dzier¿oniowski", "Zamoœæ", "tomaszowski (lubelskie)", "tarnobrzeski", "z¹bkowicki", "zawierciañski", "staszowski", "Opole", "strzelecki", "myszkowski", "k³odzki", "piñczowski", "Tarnobrzeg", "ni¿añski", "nyski", "buski", "tarnogórski", "krapkowicki", "prudnicki", "gliwicki", "bêdziñski", "mielecki", "miechowski", "kolbuszowski", "lubaczowski", "D¹browa-Górnicza", "okulski", "le¿ajski", "Bytom", "Piekary Œl¹skie", "kêdzierzyñsko-kozielski", "rzeszowski", "kazimierski", "Gliwice", "Zabrze", "d¹browski", "przeworski", "Siemianowice Œl¹skie", "Chorzów", "Ruda Œl¹ska", "proszowicki", "Œwiêtoch³owice", "Sosnowiec", "g³ubczycki", "krakowski", "Katowice", "³añcucki", _
        "jaros³awski", "Mys³owice", "raciborski", "miko³owski", "tarnowski", "chrzanowski", "brzeski (ma³opolskie)", "ropczycko-sêdziszowski", "rybnicki", "dêbicki", "Rybnik", "Tychy", "bieruñsko-lêdziñski", "bocheñski", "Kraków", "Rzeszów", "oœwiêcimski", "wielicki", "pszczyñski", "Tarnów", "wodzis³awski", "wadowicki", "przemyski", "strzy¿owski", "bielski (œl¹skie)", "cieszyñski", "myœlenicki", "jasielski", "Bielsko-Bia³a", "kroœnieñski", "brzozowski", "Przemyœl", "limanowski", "¿ywiecki", "suski", "nowos¹decki", "gorlicki", "Krosno", "bieszczadzki", "sanocki", "Nowy S¹cz", "nowotarski", "leski", "tatrzañski", "¯ory", "Jastrzêbie-Zdrój", "Jaworzno", "³ódzki wschodni")) _
        .Group.Name = "Grupa 3"
        
        ActiveSheet.Shapes.Range(Array("Grupa 1", "Grupa 2", "Grupa 3")).Group.Name = "Polska"
        
        'Wykres 2011
        With ActiveSheet.ChartObjects("Wykres 399").Chart.SeriesCollection(1)

           .Values = Worksheets("Mapka dane").Range("C" & 434 & ":E" & 434)
    
        End With
    
        'Wykres 2035
        With ActiveSheet.ChartObjects("Wykres 784").Chart.SeriesCollection(1)

            .Values = Worksheets("Mapka dane").Range("BW" & 434 & ":BY" & 434)
    
        End With
    
        'Wykres dynamiki
        With ActiveSheet.ChartObjects("Wykres 1").Chart
    
            .ChartTitle.Text = "Polska - prognoza liczby ludnoœci na lata 2011-2035"
            .SeriesCollection(1).Values = Worksheets("Mapka dane").Range("C" & 434 & "," & "F" & 434 & "," & "I" & 434 & "," & "L" & 434 & "," & "O" & 434 & "," & "R" & 434 & "," & "U" & 434 & "," & "X" & 434 & "," & "AA" & 434 & "," & "AD" & 434 & "," & "AG" & 434 & "," & "AJ" & 434 & "," & "AM" & 434 & "," & "AP" & 434 & "," & "AS" & 434 & "," & "AV" & 434 & "," & "AY" & 434 & "," & "BB" & 434 & "," & "BE" & 434 & "," & "BK" & 434 & "," & "BN" & 434 & "," & "BQ" & 434 & "," & "BT" & 434 & "," & "BW" & 434)
            .SeriesCollection(2).Values = Worksheets("Mapka dane").Range("D" & 434 & "," & "G" & 434 & "," & "J" & 434 & "," & "M" & 434 & "," & "P" & 434 & "," & "S" & 434 & "," & "V" & 434 & "," & "Y" & 434 & "," & "AB" & 434 & "," & "AE" & 434 & "," & "AH" & 434 & "," & "AK" & 434 & "," & "AN" & 434 & "," & "AQ" & 434 & "," & "AT" & 434 & "," & "AW" & 434 & "," & "AZ" & 434 & "," & "BC" & 434 & "," & "BF" & 434 & "," & "BL" & 434 & "," & "BO" & 434 & "," & "BR" & 434 & "," & "BU" & 434 & "," & "BX" & 434)
            .SeriesCollection(3).Values = Worksheets("Mapka dane").Range("E" & 434 & "," & "H" & 434 & "," & "K" & 434 & "," & "N" & 434 & "," & "Q" & 434 & "," & "T" & 434 & "," & "W" & 434 & "," & "Z" & 434 & "," & "AC" & 434 & "," & "AF" & 434 & "," & "AI" & 434 & "," & "AL" & 434 & "," & "AO" & 434 & "," & "AR" & 434 & "," & "AU" & 434 & "," & "AX" & 434 & "," & "BA" & 434 & "," & "BD" & 434 & "," & "BG" & 434 & "," & "BM" & 434 & "," & "BP" & 434 & "," & "BS" & 434 & "," & "BV" & 434 & "," & "BY" & 434)
    
        End With
               
    ElseIf ActiveSheet.Shapes.Range(Array("Woj")).Glow.Radius <> 0 Then
    
        With ActiveSheet.Shapes("pole tekstowe 2")
    
            .Visible = False
            .TextFrame.Characters.Text = ""
    
        End With
        
        counter = ""
        
        With Sheets("Powiaty").Shapes(counter)

            .Fill.Transparency = 0

        End With
    
        ActiveSheet.Shapes.Range(Array("pomorskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("zachodniopomorskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("warmiñsko-mazurskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("kujawsko-pomorskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("wielkopolskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("lubuskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("dolnoœl¹skie")).Ungroup
        ActiveSheet.Shapes.Range(Array("opolskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("œl¹skie")).Ungroup
        ActiveSheet.Shapes.Range(Array("mazowieckie")).Ungroup
        ActiveSheet.Shapes.Range(Array("podlaskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("œwiêtokrzyskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("lubelskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("³ódzkie")).Ungroup
        ActiveSheet.Shapes.Range(Array("podkarpackie")).Ungroup
        ActiveSheet.Shapes.Range(Array("ma³opolskie")).Ungroup
        
        ActiveSheet.Shapes.Range(Array("Powiat")).Glow.Radius = 0
        ActiveSheet.Shapes.Range(Array("Woj")).Glow.Radius = 0
    
        With ActiveSheet.Shapes.Range(Array("Kraj")).Glow
            .color.ObjectThemeColor = msoThemeColorBackground1
            ' .Transparency = 0.2
            .Radius = 5
        End With
        
        With ActiveSheet.Shapes.Range(Array( _
        "hajnowski", "Toruñ", "miêdzyrzecki", "sulêciñski", "pucki", "wejherowski", "lêborski", "s³upski", "s³awieñski", "Gdynia", "S³upsk", "kartuski", "suwalski", "bytowski", "Gdañsk", "braniewski", "go³dapski", "bartoszycki", "wêgorzewski", "sejneñski", "kêtrzyñski", "koszaliñski", "nowodworski (pomorskie)", _
        "gdañski", "elbl¹ski", "Elbl¹g", "olecki", "ko³obrzeski", "Koszalin", "lidzbarski", "Suwa³ki", "gryficki", "gi¿ycki", "koœcierski", "malborski", "bia³ogardzki", "tczewski", "starogardzki", "kamieñski", "olsztyñski", "augustowski", "Œwinoujœcie", "ostródzki", "e³cki", "chojnicki", "sztumski", "szczecinecki", "cz³uchowski", "œwidwiñski", "mr¹gowski", _
        "piski", "³obeski", "i³awski", "kwidzyñski", "goleniowski", "Olsztyn", "grajewski", "policki", "szczycieñski", "sokólski", "tucholski", "moniecki", "drawski", "œwiecki", "z³otowski", "sêpoleñski", "grudzi¹dzki", "Szczecin", "stargardzki", "kolneñski", "nowomiejski", "Grudzi¹dz", "nidzicki", "ostro³êcki", "³om¿yñski", "wa³ecki", "bia³ostocki", "brodnicki", "bydgoski", "dzia³dowski", "gryfiñski", "che³miñski", "w¹brzeski", "choszczeñski", _
        "przasnyski", "pilski", "pyrzycki", "nakielski", "m³awski", "toruñski", "Bia³ystok", "golubsko-dobrzyñski", "zambrowski", "£om¿a", "rypiñski", "Bydgoszcz", "¿uromiñski", "wysokomazowiecki", "strzelecko-drezdenecki", "czarnkowsko-trzcianecki", "makowski", "Ostro³êka", "ciechanowski", "myœliborski", "chodzieski")) _
        .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    
         With ActiveSheet.Shapes.Range(Array( _
        "hajnowski", "Toruñ", "miêdzyrzecki", "sulêciñski", "pucki", "wejherowski", "lêborski", "s³upski", "s³awieñski", "Gdynia", "S³upsk", "kartuski", "suwalski", "bytowski", "Gdañsk", "braniewski", "go³dapski", "bartoszycki", "wêgorzewski", "sejneñski", "kêtrzyñski", "koszaliñski", "nowodworski (pomorskie)", _
        "gdañski", "elbl¹ski", "Elbl¹g", "olecki", "ko³obrzeski", "Koszalin", "lidzbarski", "Suwa³ki", "gryficki", "gi¿ycki", "koœcierski", "malborski", "bia³ogardzki", "tczewski", "starogardzki", "kamieñski", "olsztyñski", "augustowski", "Œwinoujœcie", "ostródzki", "e³cki", "chojnicki", "sztumski", "szczecinecki", "cz³uchowski", "œwidwiñski", "mr¹gowski", _
        "piski", "³obeski", "i³awski", "kwidzyñski", "goleniowski", "Olsztyn", "grajewski", "policki", "szczycieñski", "sokólski", "tucholski", "moniecki", "drawski", "œwiecki", "z³otowski", "sêpoleñski", "grudzi¹dzki", "Szczecin", "stargardzki", "kolneñski", "nowomiejski", "Grudzi¹dz", "nidzicki", "ostro³êcki", "³om¿yñski", "wa³ecki", "bia³ostocki", "brodnicki", "bydgoski", "dzia³dowski", "gryfiñski", "che³miñski", "w¹brzeski", "choszczeñski", _
        "przasnyski", "pilski", "pyrzycki", "nakielski", "m³awski", "toruñski", "Bia³ystok", "golubsko-dobrzyñski", "zambrowski", "£om¿a", "rypiñski", "Bydgoszcz", "¿uromiñski", "wysokomazowiecki", "strzelecko-drezdenecki", "czarnkowsko-trzcianecki", "makowski", "Ostro³êka", "ciechanowski", "myœliborski", "chodzieski")) _
        .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
            .Visible = msoTrue
            .Weight = 2.5
        End With
        
         ActiveSheet.Shapes.Range(Array( _
        "hajnowski", "Toruñ", "miêdzyrzecki", "sulêciñski", "pucki", "wejherowski", "lêborski", "s³upski", "s³awieñski", "Gdynia", "S³upsk", "kartuski", "suwalski", "bytowski", "Gdañsk", "braniewski", "go³dapski", "bartoszycki", "wêgorzewski", "sejneñski", "kêtrzyñski", "koszaliñski", "nowodworski (pomorskie)", _
        "gdañski", "elbl¹ski", "Elbl¹g", "olecki", "ko³obrzeski", "Koszalin", "lidzbarski", "Suwa³ki", "gryficki", "gi¿ycki", "koœcierski", "malborski", "bia³ogardzki", "tczewski", "starogardzki", "kamieñski", "olsztyñski", "augustowski", "Œwinoujœcie", "ostródzki", "e³cki", "chojnicki", "sztumski", "szczecinecki", "cz³uchowski", "œwidwiñski", "mr¹gowski", _
        "piski", "³obeski", "i³awski", "kwidzyñski", "goleniowski", "Olsztyn", "grajewski", "policki", "szczycieñski", "sokólski", "tucholski", "moniecki", "drawski", "œwiecki", "z³otowski", "sêpoleñski", "grudzi¹dzki", "Szczecin", "stargardzki", "kolneñski", "nowomiejski", "Grudzi¹dz", "nidzicki", "ostro³êcki", "³om¿yñski", "wa³ecki", "bia³ostocki", "brodnicki", "bydgoski", "dzia³dowski", "gryfiñski", "che³miñski", "w¹brzeski", "choszczeñski", _
        "przasnyski", "pilski", "pyrzycki", "nakielski", "m³awski", "toruñski", "Bia³ystok", "golubsko-dobrzyñski", "zambrowski", "£om¿a", "rypiñski", "Bydgoszcz", "¿uromiñski", "wysokomazowiecki", "strzelecko-drezdenecki", "czarnkowsko-trzcianecki", "makowski", "Ostro³êka", "ciechanowski", "myœliborski", "chodzieski")) _
        .Group.Name = "Grupa 1"
        
         With ActiveSheet.Shapes.Range(Array("¿niñski", "w¹growiecki", "Konin", "Sopot", "Siedlce", "ostrowski (mazowieckie)", "lipnowski", "bielski (podlaskie)", "inowroc³awski", "sierpecki", "gorzowski", "aleksandrowski", "p³oñski", "obornicki", "wyszkowski", "miêdzychodzki", "Gorzów Wielkopolski", "mogileñski", "pu³tuski", "szamotulski", "p³ocki", "w³oc³awski", "siemiatycki", "gnieŸnieñski", "W³oc³awek", "soko³owski", "wêgrowski", "radziejowski", "poznañski", "nowodworski (mazowieckie)", "s³ubicki", "P³ock", "wo³omiñski", "legionowski", "s³upecki", "nowotomyski", "koniñski", "Poznañ", "gostyniñski", "siedlecki", "³osicki", "œwiebodziñski", "wrzesiñski", "kolski", "miñski", "warszawski zachodni", "kutnowski", "sochaczewski", "Warszawa")) _
         .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    
         With ActiveSheet.Shapes.Range(Array("¿niñski", "w¹growiecki", "Konin", "Sopot", "Siedlce", "ostrowski (mazowieckie)", "lipnowski", "bielski (podlaskie)", "inowroc³awski", "sierpecki", "gorzowski", "aleksandrowski", "p³oñski", "obornicki", "wyszkowski", "miêdzychodzki", "Gorzów Wielkopolski", "mogileñski", "pu³tuski", "szamotulski", "p³ocki", "w³oc³awski", "siemiatycki", "gnieŸnieñski", "W³oc³awek", "soko³owski", "wêgrowski", "radziejowski", "poznañski", "nowodworski (mazowieckie)", "s³ubicki", "P³ock", "wo³omiñski", "legionowski", "s³upecki", "nowotomyski", "koniñski", "Poznañ", "gostyniñski", "siedlecki", "³osicki", "œwiebodziñski", "wrzesiñski", "kolski", "miñski", "warszawski zachodni", "kutnowski", "sochaczewski", "Warszawa")) _
         .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
            .Visible = msoTrue
            .Weight = 2.5
        End With
        
        ActiveSheet.Shapes.Range(Array("¿niñski", "w¹growiecki", "Konin", "Sopot", "Siedlce", "ostrowski (mazowieckie)", "lipnowski", "bielski (podlaskie)", "inowroc³awski", "sierpecki", "gorzowski", "aleksandrowski", "p³oñski", "obornicki", "wyszkowski", "miêdzychodzki", "Gorzów Wielkopolski", "mogileñski", "pu³tuski", "szamotulski", "p³ocki", "w³oc³awski", "siemiatycki", "gnieŸnieñski", "W³oc³awek", "soko³owski", "wêgrowski", "radziejowski", "poznañski", "nowodworski (mazowieckie)", "s³ubicki", "P³ock", "wo³omiñski", "legionowski", "s³upecki", "nowotomyski", "koniñski", "Poznañ", "gostyniñski", "siedlecki", "³osicki", "œwiebodziñski", "wrzesiñski", "kolski", "miñski", "warszawski zachodni", "kutnowski", "sochaczewski", "Warszawa")) _
         .Group.Name = "Grupa 2"
        
       With ActiveSheet.Shapes.Range(Array( _
        "bialski", "œredzki (wielkopolskie)", "grodziski (wielkopolskie)", "³owicki", "kroœnieñski (lubuskie)", "wolsztyñski", "zielonogórski", "otwocki", "koœciañski", "pruszkowski", "³êczycki", "grodziski (mazowieckie)", "œremski", "turecki", "¿yrardowski", "piaseczyñski", "skierniewicki", "jarociñski", "Bia³a Podlaska", "³ukowski", "pleszewski", "poddêbicki", "nowosolski", "garwoliñski", "kaliski", "Zielona Góra", "wschowski", "zgierski", "radzyñski", "leszczyñski", "Skierniewice", "gostyñski", "grójecki", "brzeziñski", "¿arski", "krotoszyñski", "rawski", "Leszno", "kozienicki", "w³odawski", "£ódŸ", "parczewski", "sieradzki", "rycki", "g³ogowski", "¿agañski", "ostrowski (wielkopolskie)", "pabianicki", "górowski", "Kalisz", "tomaszowski (³ódzkie)", "zduñskowolski", "³aski", "bia³obrzeski", "lubartowski", "rawicki", "polkowicki", _
        "milicki", "piotrkowski", "przysuski", "lubiñski", "radomski", "pu³awski", "opoczyñski", "trzebnicki", "boles³awiecki", "ostrzeszowski", "be³chatowski", "wo³owski", "zgorzelecki", "Radom", "zwoleñski", "lubelski", "³êczyñski", "Piotrków Trybunalski", "che³mski", "oleœnicki", "wieruszowski", "legnicki", "wieluñski", "kêpiñski", "szyd³owiecki", "Lublin", "konecki", "opolski (lubelskie)", "œwidnicki (lubelskie)", "lipski", "z³otoryjski", "pajêczañski", "radomszczañski", "œredzki (dolnyœl¹sk)", "wroc³awski", "Legnica", "Che³m", "lwówecki", "lubañski", "starachowicki", "Wroc³aw", "skar¿yski", "namys³owski", "krasnostawski", "kluczborski", "jaworski", "oleski", "opatowski", "o³awski", "ostrowiecki", "kraœnicki", "k³obucki", "hrubieszowski", "kielecki", "jeleniogórski", "œwidnicki (dolnoœl¹skie)", "czêstochowski", "w³oszczowski", "brzeski (opolskie)", "zamojski", _
        "Jelenia Góra", "opolski (opolskie)", "janowski", "bi³gorajski", "Kielce", "kamiennogórski", "strzeliñski", "sandomierski", "wa³brzyski", "Czêstochowa", "jêdrzejowski", "lubliniecki", "stalowowolski", "dzier¿oniowski", "Zamoœæ", "tomaszowski (lubelskie)", "tarnobrzeski", "z¹bkowicki", "zawierciañski", "staszowski", "Opole", "strzelecki", "myszkowski", "k³odzki", "piñczowski", "Tarnobrzeg", "ni¿añski", "nyski", "buski", "tarnogórski", "krapkowicki", "prudnicki", "gliwicki", "bêdziñski", "mielecki", "miechowski", "kolbuszowski", "lubaczowski", "D¹browa-Górnicza", "okulski", "le¿ajski", "Bytom", "Piekary Œl¹skie", "kêdzierzyñsko-kozielski", "rzeszowski", "kazimierski", "Gliwice", "Zabrze", "d¹browski", "przeworski", "Siemianowice Œl¹skie", "Chorzów", "Ruda Œl¹ska", "proszowicki", "Œwiêtoch³owice", "Sosnowiec", "g³ubczycki", "krakowski", "Katowice", "³añcucki", _
        "jaros³awski", "Mys³owice", "raciborski", "miko³owski", "tarnowski", "chrzanowski", "brzeski (ma³opolskie)", "ropczycko-sêdziszowski", "rybnicki", "dêbicki", "Rybnik", "Tychy", "bieruñsko-lêdziñski", "bocheñski", "Kraków", "Rzeszów", "oœwiêcimski", "wielicki", "pszczyñski", "Tarnów", "wodzis³awski", "wadowicki", "przemyski", "strzy¿owski", "bielski (œl¹skie)", "cieszyñski", "myœlenicki", "jasielski", "Bielsko-Bia³a", "kroœnieñski", "brzozowski", "Przemyœl", "limanowski", "¿ywiecki", "suski", "nowos¹decki", "gorlicki", "Krosno", "bieszczadzki", "sanocki", "Nowy S¹cz", "nowotarski", "leski", "tatrzañski", "¯ory", "Jastrzêbie-Zdrój", "Jaworzno", "³ódzki wschodni")) _
        .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(220, 20, 60)
            .Solid
        End With
    
       With ActiveSheet.Shapes.Range(Array( _
        "bialski", "œredzki (wielkopolskie)", "grodziski (wielkopolskie)", "³owicki", "kroœnieñski (lubuskie)", "wolsztyñski", "zielonogórski", "otwocki", "koœciañski", "pruszkowski", "³êczycki", "grodziski (mazowieckie)", "œremski", "turecki", "¿yrardowski", "piaseczyñski", "skierniewicki", "jarociñski", "Bia³a Podlaska", "³ukowski", "pleszewski", "poddêbicki", "nowosolski", "garwoliñski", "kaliski", "Zielona Góra", "wschowski", "zgierski", "radzyñski", "leszczyñski", "Skierniewice", "gostyñski", "grójecki", "brzeziñski", "¿arski", "krotoszyñski", "rawski", "Leszno", "kozienicki", "w³odawski", "£ódŸ", "parczewski", "sieradzki", "rycki", "g³ogowski", "¿agañski", "ostrowski (wielkopolskie)", "pabianicki", "górowski", "Kalisz", "tomaszowski (³ódzkie)", "zduñskowolski", "³aski", "bia³obrzeski", "lubartowski", "rawicki", "polkowicki", _
        "milicki", "piotrkowski", "przysuski", "lubiñski", "radomski", "pu³awski", "opoczyñski", "trzebnicki", "boles³awiecki", "ostrzeszowski", "be³chatowski", "wo³owski", "zgorzelecki", "Radom", "zwoleñski", "lubelski", "³êczyñski", "Piotrków Trybunalski", "che³mski", "oleœnicki", "wieruszowski", "legnicki", "wieluñski", "kêpiñski", "szyd³owiecki", "Lublin", "konecki", "opolski (lubelskie)", "œwidnicki (lubelskie)", "lipski", "z³otoryjski", "pajêczañski", "radomszczañski", "œredzki (dolnyœl¹sk)", "wroc³awski", "Legnica", "Che³m", "lwówecki", "lubañski", "starachowicki", "Wroc³aw", "skar¿yski", "namys³owski", "krasnostawski", "kluczborski", "jaworski", "oleski", "opatowski", "o³awski", "ostrowiecki", "kraœnicki", "k³obucki", "hrubieszowski", "kielecki", "jeleniogórski", "œwidnicki (dolnoœl¹skie)", "czêstochowski", "w³oszczowski", "brzeski (opolskie)", "zamojski", _
        "Jelenia Góra", "opolski (opolskie)", "janowski", "bi³gorajski", "Kielce", "kamiennogórski", "strzeliñski", "sandomierski", "wa³brzyski", "Czêstochowa", "jêdrzejowski", "lubliniecki", "stalowowolski", "dzier¿oniowski", "Zamoœæ", "tomaszowski (lubelskie)", "tarnobrzeski", "z¹bkowicki", "zawierciañski", "staszowski", "Opole", "strzelecki", "myszkowski", "k³odzki", "piñczowski", "Tarnobrzeg", "ni¿añski", "nyski", "buski", "tarnogórski", "krapkowicki", "prudnicki", "gliwicki", "bêdziñski", "mielecki", "miechowski", "kolbuszowski", "lubaczowski", "D¹browa-Górnicza", "okulski", "le¿ajski", "Bytom", "Piekary Œl¹skie", "kêdzierzyñsko-kozielski", "rzeszowski", "kazimierski", "Gliwice", "Zabrze", "d¹browski", "przeworski", "Siemianowice Œl¹skie", "Chorzów", "Ruda Œl¹ska", "proszowicki", "Œwiêtoch³owice", "Sosnowiec", "g³ubczycki", "krakowski", "Katowice", "³añcucki", _
        "jaros³awski", "Mys³owice", "raciborski", "miko³owski", "tarnowski", "chrzanowski", "brzeski (ma³opolskie)", "ropczycko-sêdziszowski", "rybnicki", "dêbicki", "Rybnik", "Tychy", "bieruñsko-lêdziñski", "bocheñski", "Kraków", "Rzeszów", "oœwiêcimski", "wielicki", "pszczyñski", "Tarnów", "wodzis³awski", "wadowicki", "przemyski", "strzy¿owski", "bielski (œl¹skie)", "cieszyñski", "myœlenicki", "jasielski", "Bielsko-Bia³a", "kroœnieñski", "brzozowski", "Przemyœl", "limanowski", "¿ywiecki", "suski", "nowos¹decki", "gorlicki", "Krosno", "bieszczadzki", "sanocki", "Nowy S¹cz", "nowotarski", "leski", "tatrzañski", "¯ory", "Jastrzêbie-Zdrój", "Jaworzno", "³ódzki wschodni")) _
        .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(220, 20, 60)
            .Visible = msoTrue
            .Weight = 2.5
        End With
        
        ActiveSheet.Shapes.Range(Array( _
        "bialski", "œredzki (wielkopolskie)", "grodziski (wielkopolskie)", "³owicki", "kroœnieñski (lubuskie)", "wolsztyñski", "zielonogórski", "otwocki", "koœciañski", "pruszkowski", "³êczycki", "grodziski (mazowieckie)", "œremski", "turecki", "¿yrardowski", "piaseczyñski", "skierniewicki", "jarociñski", "Bia³a Podlaska", "³ukowski", "pleszewski", "poddêbicki", "nowosolski", "garwoliñski", "kaliski", "Zielona Góra", "wschowski", "zgierski", "radzyñski", "leszczyñski", "Skierniewice", "gostyñski", "grójecki", "brzeziñski", "¿arski", "krotoszyñski", "rawski", "Leszno", "kozienicki", "w³odawski", "£ódŸ", "parczewski", "sieradzki", "rycki", "g³ogowski", "¿agañski", "ostrowski (wielkopolskie)", "pabianicki", "górowski", "Kalisz", "tomaszowski (³ódzkie)", "zduñskowolski", "³aski", "bia³obrzeski", "lubartowski", "rawicki", "polkowicki", _
        "milicki", "piotrkowski", "przysuski", "lubiñski", "radomski", "pu³awski", "opoczyñski", "trzebnicki", "boles³awiecki", "ostrzeszowski", "be³chatowski", "wo³owski", "zgorzelecki", "Radom", "zwoleñski", "lubelski", "³êczyñski", "Piotrków Trybunalski", "che³mski", "oleœnicki", "wieruszowski", "legnicki", "wieluñski", "kêpiñski", "szyd³owiecki", "Lublin", "konecki", "opolski (lubelskie)", "œwidnicki (lubelskie)", "lipski", "z³otoryjski", "pajêczañski", "radomszczañski", "œredzki (dolnyœl¹sk)", "wroc³awski", "Legnica", "Che³m", "lwówecki", "lubañski", "starachowicki", "Wroc³aw", "skar¿yski", "namys³owski", "krasnostawski", "kluczborski", "jaworski", "oleski", "opatowski", "o³awski", "ostrowiecki", "kraœnicki", "k³obucki", "hrubieszowski", "kielecki", "jeleniogórski", "œwidnicki (dolnoœl¹skie)", "czêstochowski", "w³oszczowski", "brzeski (opolskie)", "zamojski", _
        "Jelenia Góra", "opolski (opolskie)", "janowski", "bi³gorajski", "Kielce", "kamiennogórski", "strzeliñski", "sandomierski", "wa³brzyski", "Czêstochowa", "jêdrzejowski", "lubliniecki", "stalowowolski", "dzier¿oniowski", "Zamoœæ", "tomaszowski (lubelskie)", "tarnobrzeski", "z¹bkowicki", "zawierciañski", "staszowski", "Opole", "strzelecki", "myszkowski", "k³odzki", "piñczowski", "Tarnobrzeg", "ni¿añski", "nyski", "buski", "tarnogórski", "krapkowicki", "prudnicki", "gliwicki", "bêdziñski", "mielecki", "miechowski", "kolbuszowski", "lubaczowski", "D¹browa-Górnicza", "okulski", "le¿ajski", "Bytom", "Piekary Œl¹skie", "kêdzierzyñsko-kozielski", "rzeszowski", "kazimierski", "Gliwice", "Zabrze", "d¹browski", "przeworski", "Siemianowice Œl¹skie", "Chorzów", "Ruda Œl¹ska", "proszowicki", "Œwiêtoch³owice", "Sosnowiec", "g³ubczycki", "krakowski", "Katowice", "³añcucki", _
        "jaros³awski", "Mys³owice", "raciborski", "miko³owski", "tarnowski", "chrzanowski", "brzeski (ma³opolskie)", "ropczycko-sêdziszowski", "rybnicki", "dêbicki", "Rybnik", "Tychy", "bieruñsko-lêdziñski", "bocheñski", "Kraków", "Rzeszów", "oœwiêcimski", "wielicki", "pszczyñski", "Tarnów", "wodzis³awski", "wadowicki", "przemyski", "strzy¿owski", "bielski (œl¹skie)", "cieszyñski", "myœlenicki", "jasielski", "Bielsko-Bia³a", "kroœnieñski", "brzozowski", "Przemyœl", "limanowski", "¿ywiecki", "suski", "nowos¹decki", "gorlicki", "Krosno", "bieszczadzki", "sanocki", "Nowy S¹cz", "nowotarski", "leski", "tatrzañski", "¯ory", "Jastrzêbie-Zdrój", "Jaworzno", "³ódzki wschodni")) _
        .Group.Name = "Grupa 3"
        
        ActiveSheet.Shapes.Range(Array("Grupa 1", "Grupa 2", "Grupa 3")).Group.Name = "Polska"
           
        'Wykres 2011
        With ActiveSheet.ChartObjects("Wykres 399")
        
            With .Chart.SeriesCollection(1)

               .Values = Worksheets("Mapka dane").Range("C" & 434 & ":E" & 434)
            
            End With
        
        End With
    
        'Wykres 2035
        With ActiveSheet.ChartObjects("Wykres 784")
        
            With .Chart.SeriesCollection(1)

                .Values = Worksheets("Mapka dane").Range("BW" & 434 & ":BY" & 434)
            
            End With
    
        End With
    
        'Wykres dynamiki
        With ActiveSheet.ChartObjects("Wykres 1").Chart
    
            .ChartTitle.Text = "Polska - prognoza liczby ludnoœci na lata 2011-2035"
            .SeriesCollection(1).Values = Worksheets("Mapka dane").Range("C" & 434 & "," & "F" & 434 & "," & "I" & 434 & "," & "L" & 434 & "," & "O" & 434 & "," & "R" & 434 & "," & "U" & 434 & "," & "X" & 434 & "," & "AA" & 434 & "," & "AD" & 434 & "," & "AG" & 434 & "," & "AJ" & 434 & "," & "AM" & 434 & "," & "AP" & 434 & "," & "AS" & 434 & "," & "AV" & 434 & "," & "AY" & 434 & "," & "BB" & 434 & "," & "BE" & 434 & "," & "BK" & 434 & "," & "BN" & 434 & "," & "BQ" & 434 & "," & "BT" & 434 & "," & "BW" & 434)
            .SeriesCollection(2).Values = Worksheets("Mapka dane").Range("D" & 434 & "," & "G" & 434 & "," & "J" & 434 & "," & "M" & 434 & "," & "P" & 434 & "," & "S" & 434 & "," & "V" & 434 & "," & "Y" & 434 & "," & "AB" & 434 & "," & "AE" & 434 & "," & "AH" & 434 & "," & "AK" & 434 & "," & "AN" & 434 & "," & "AQ" & 434 & "," & "AT" & 434 & "," & "AW" & 434 & "," & "AZ" & 434 & "," & "BC" & 434 & "," & "BF" & 434 & "," & "BL" & 434 & "," & "BO" & 434 & "," & "BR" & 434 & "," & "BU" & 434 & "," & "BX" & 434)
            .SeriesCollection(3).Values = Worksheets("Mapka dane").Range("E" & 434 & "," & "H" & 434 & "," & "K" & 434 & "," & "N" & 434 & "," & "Q" & 434 & "," & "T" & 434 & "," & "W" & 434 & "," & "Z" & 434 & "," & "AC" & 434 & "," & "AF" & 434 & "," & "AI" & 434 & "," & "AL" & 434 & "," & "AO" & 434 & "," & "AR" & 434 & "," & "AU" & 434 & "," & "AX" & 434 & "," & "BA" & 434 & "," & "BD" & 434 & "," & "BG" & 434 & "," & "BM" & 434 & "," & "BP" & 434 & "," & "BS" & 434 & "," & "BV" & 434 & "," & "BY" & 434)
    
        End With
    
    End If
    
End Sub
Public Sub CalyPowiat()

    If ActiveSheet.Shapes.Range(Array("Powiat")).Glow.Radius <> 0 Then Exit Sub
    
    If ActiveSheet.Shapes.Range(Array("Kraj")).Glow.Radius <> 0 Then

        ActiveSheet.Shapes.Range(Array("Polska")).Ungroup
        ActiveSheet.Shapes.Range(Array("Grupa 1")).Ungroup
        ActiveSheet.Shapes.Range(Array("Grupa 2")).Ungroup
        ActiveSheet.Shapes.Range(Array("Grupa 3")).Ungroup
    
        ActiveSheet.Shapes.Range(Array("Kraj")).Glow.Radius = 0
        ActiveSheet.Shapes.Range(Array("Woj")).Glow.Radius = 0
    
        With ActiveSheet.Shapes.Range(Array("Powiat")).Glow
            .color.ObjectThemeColor = msoThemeColorBackground1
            ' .Transparency = 0.2
            .Radius = 5
        End With
      
        Range("H42:K43").FormulaR1C1 = "GOP"
        Range("A44:S55").Borders(xlEdgeLeft).LineStyle = xlContinuous
        Range("A44:S55").Borders(xlEdgeTop).LineStyle = xlContinuous
        Range("A44:S55").Borders(xlEdgeRight).LineStyle = xlContinuous
           
        ActiveSheet.Shapes.Range(Array(" ")).Visible = msoTrue
           
        ActiveSheet.Shapes.Range(Array(" ")).Ungroup
       
        With ActiveSheet.Shapes.Range(Array("pucki", "wejherowski", "lêborski", "s³upski" _
            , "bytowski", "cz³uchowski", "chojnicki", "koœcierski", "kartuski", "gdañski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzyñski")).Fill
            
            .ForeColor.RGB = RGB(255, 192, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("pucki", "wejherowski", "lêborski", "s³upski" _
            , "bytowski", "cz³uchowski", "chojnicki", "koœcierski", "kartuski", "gdañski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzyñski")).Line
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("s³awieñski", "bia³ogardzki", "koszaliñski", _
            "szczecinecki", "wa³ecki", "drawski", "œwidwiñski", "ko³obrzeski", "gryficki", _
            "³obeski", "Œwinoujœcie", "goleniowski", "stargardzki", "choszczeñski", _
            "pyrzycki", "myœliborski", "gryfiñski", "policki", "kamieñski")).Fill
            
            .ForeColor.RGB = RGB(192, 0, 0)
        End With

        With ActiveSheet.Shapes.Range(Array("s³awieñski", "bia³ogardzki", "koszaliñski", _
            "szczecinecki", "wa³ecki", "drawski", "œwidwiñski", "ko³obrzeski", "gryficki", _
            "³obeski", "Œwinoujœcie", "goleniowski", "stargardzki", "choszczeñski", _
            "pyrzycki", "myœliborski", "gryfiñski", "policki", "kamieñski")).Line
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("braniewski", "elbl¹ski", "i³awski", _
            "nowomiejski", "dzia³dowski", "ostródzki", "lidzbarski", "bartoszycki", _
            "olsztyñski", "nidzicki", "szczycieñski", "mr¹gowski", "kêtrzyñski", _
            "wêgorzewski", "go³dapski", "olecki", "gi¿ycki", "e³cki", "piski")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent3
        End With
        
        With ActiveSheet.Shapes.Range(Array("braniewski", "elbl¹ski", "i³awski", _
            "nowomiejski", "dzia³dowski", "ostródzki", "lidzbarski", "bartoszycki", _
            "olsztyñski", "nidzicki", "szczycieñski", "mr¹gowski", "kêtrzyñski", _
            "wêgorzewski", "go³dapski", "olecki", "gi¿ycki", "e³cki", "piski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("sêpoleñski", "tucholski", "che³miñski", _
            "œwiecki", "grudzi¹dzki", "w¹brzeski", "brodnicki", "rypiñski", _
            "golubsko-dobrzyñski", "lipnowski", "w³oc³awski", "radziejowski", _
            "aleksandrowski", "toruñski", "inowroc³awski", "mogileñski", "bydgoski", _
            "¿niñski")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent4
        End With
        
        With ActiveSheet.Shapes.Range(Array("sêpoleñski", "tucholski", "che³miñski", _
            "œwiecki", "grudzi¹dzki", "w¹brzeski", "brodnicki", "rypiñski", _
            "golubsko-dobrzyñski", "lipnowski", "w³oc³awski", "radziejowski", _
            "aleksandrowski", "toruñski", "inowroc³awski", "mogileñski", "bydgoski", _
            "¿niñski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("z³otowski", "pilski", "nakielski", _
            "chodzieski", "w¹growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "miêdzychodzki", "nowotomyski", "wolsztyñski", "grodziski (wielkopolskie)", _
            "poznañski", "gnieŸnieñski", "kolski", "koniñski", "s³upecki", "wrzesiñski", _
            "œremski", "jarociñski", "œredzki (wielkopolskie)", "leszczyñski", "gostyñski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszyñski", "ostrzeszowski", "kêpiñski", "turecki", "koœciañski")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent5
        End With
        
        With ActiveSheet.Shapes.Range(Array("z³otowski", "pilski", "nakielski", _
            "chodzieski", "w¹growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "miêdzychodzki", "nowotomyski", "wolsztyñski", "grodziski (wielkopolskie)", _
            "poznañski", "gnieŸnieñski", "kolski", "koniñski", "s³upecki", "wrzesiñski", _
            "œremski", "jarociñski", "œredzki (wielkopolskie)", "leszczyñski", "gostyñski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszyñski", "ostrzeszowski", "kêpiñski", "turecki", "koœciañski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
     
        With ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "gorzowski", _
            "sulêciñski", "s³ubicki", "miêdzyrzecki", "œwiebodziñski", "zielonogórski", _
            "kroœnieñski (lubuskie)", "¿arski", "nowosolski", "¿agañski", "wschowski")). _
            Fill
            
            .ForeColor.RGB = RGB(165, 27, 96)
        End With
        
        With ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "gorzowski", _
            "sulêciñski", "s³ubicki", "miêdzyrzecki", "œwiebodziñski", "zielonogórski", _
            "kroœnieñski (lubuskie)", "¿arski", "nowosolski", "¿agañski", "wschowski")). _
            Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("wa³brzyski", "kamiennogórski", "z³otoryjski" _
            , "lubañski", "zgorzelecki", "boles³awiecki", "polkowicki", "g³ogowski", _
            "górowski", "wo³owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "oleœnicki", "wroc³awski", "o³awski", "strzeliñski", "dzier¿oniowski", _
            "z¹bkowicki", "lubiñski", "œredzki (dolnyœl¹sk)", "œwidnicki (dolnoœl¹skie)", _
            "lwówecki", "jeleniogórski", "k³odzki")).Fill
            
            .ForeColor.RGB = RGB(255, 0, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("wa³brzyski", "kamiennogórski", "z³otoryjski" _
            , "lubañski", "zgorzelecki", "boles³awiecki", "polkowicki", "g³ogowski", _
            "górowski", "wo³owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "oleœnicki", "wroc³awski", "o³awski", "strzeliñski", "dzier¿oniowski", _
            "z¹bkowicki", "lubiñski", "œredzki (dolnyœl¹sk)", "œwidnicki (dolnoœl¹skie)", _
            "lwówecki", "jeleniogórski", "k³odzki")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("suwalski", "sejneñski", "augustowski", _
            "grajewski", "kolneñski", "moniecki", "sokólski", "bia³ostocki", "³om¿yñski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        End With
        
        With ActiveSheet.Shapes.Range(Array("suwalski", "sejneñski", "augustowski", _
            "grajewski", "kolneñski", "moniecki", "sokólski", "bia³ostocki", "³om¿yñski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("ostro³êcki", "przasnyski", "ciechanowski", _
            "m³awski", "¿uromiñski", "sierpecki", "gostyniñski", "p³ocki", "p³oñski", _
            "pu³tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko³owski", "³osicki", "siedlecki", "wêgrowski", _
            "wo³omiñski", "miñski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "¿yrardowski", "piaseczyñski", "otwocki", "grójecki", "przysuski", _
            "bia³obrzeski", "kozienicki", "garwoliñski", "zwoleñski", "szyd³owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Fill
            
            .ForeColor.RGB = RGB(255, 204, 51)
        End With
        
        With ActiveSheet.Shapes.Range(Array("ostro³êcki", "przasnyski", "ciechanowski", _
            "m³awski", "¿uromiñski", "sierpecki", "gostyniñski", "p³ocki", "p³oñski", _
            "pu³tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko³owski", "³osicki", "siedlecki", "wêgrowski", _
            "wo³omiñski", "miñski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "¿yrardowski", "piaseczyñski", "otwocki", "grójecki", "przysuski", _
            "bia³obrzeski", "kozienicki", "garwoliñski", "zwoleñski", "szyd³owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("rawski", "skierniewicki", "³owicki", _
            "kutnowski", "³êczycki", "zgierski", "brzeziñski", "tomaszowski (³ódzkie)", _
            "opoczyñski", "³ódzki wschodni", "piotrkowski", "radomszczañski", "pajêczañski" _
            , "be³chatowski", "pabianicki", "³aski", "zduñskowolski", "poddêbicki", _
            "sieradzki", "wieluñski", "wieruszowski")).Fill
            
            .ForeColor.RGB = RGB(0, 176, 80)
        End With
        
        With ActiveSheet.Shapes.Range(Array("rawski", "skierniewicki", "³owicki", _
            "kutnowski", "³êczycki", "zgierski", "brzeziñski", "tomaszowski (³ódzkie)", _
            "opoczyñski", "³ódzki wschodni", "piotrkowski", "radomszczañski", "pajêczañski" _
            , "be³chatowski", "pabianicki", "³aski", "zduñskowolski", "poddêbicki", _
            "sieradzki", "wieluñski", "wieruszowski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("g³ubczycki", "kêdzierzyñsko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys³owski", "brzeski (opolskie)", "nyski")).Fill
            
            .ForeColor.RGB = RGB(0, 0, 204)
        End With
        
        With ActiveSheet.Shapes.Range(Array("g³ubczycki", "kêdzierzyñsko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys³owski", "brzeski (opolskie)", "nyski")).Line
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("bieruñsko-lêdziñski", "pszczyñski", _
            "bielski (œl¹skie)", "cieszyñski", "wodzis³awski", "rybnicki", "raciborski", _
            "gliwicki", "miko³owski", "bêdziñski", "tarnogórski", "zawierciañski", _
            "myszkowski", "czêstochowski", "lubliniecki", "k³obucki", "¿ywiecki")).Fill
            
            .ForeColor.RGB = RGB(204, 153, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("bieruñsko-lêdziñski", "pszczyñski", _
            "bielski (œl¹skie)", "cieszyñski", "wodzis³awski", "rybnicki", "raciborski", _
            "gliwicki", "miko³owski", "bêdziñski", "tarnogórski", "zawierciañski", _
            "myszkowski", "czêstochowski", "lubliniecki", "k³obucki", "¿ywiecki")).Line
            
             .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("konecki", "skar¿yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w³oszczowski", "jêdrzejowski", "piñczowski", "buski", "kazimierski")).Fill
            
        .ForeColor.RGB = RGB(204, 51, 255)

        End With
        
        With ActiveSheet.Shapes.Range(Array("konecki", "skar¿yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w³oszczowski", "jêdrzejowski", "piñczowski", "buski", "kazimierski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("bialski", "radzyñski", "³ukowski", "rycki", _
            "parczewski", "w³odawski", "lubartowski", "pu³awski", "opolski (lubelskie)", _
            "lubelski", "kraœnicki", "œwidnicki (lubelskie)", "che³mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi³gorajski", "³êczyñski", _
            "tomaszowski (lubelskie)")).Fill
            
            .ForeColor.RGB = RGB(102, 0, 204)
        End With
        
        With ActiveSheet.Shapes.Range(Array("bialski", "radzyñski", "³ukowski", "rycki", _
            "parczewski", "w³odawski", "lubartowski", "pu³awski", "opolski (lubelskie)", _
            "lubelski", "kraœnicki", "œwidnicki (lubelskie)", "che³mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi³gorajski", "³êczyñski", _
            "tomaszowski (lubelskie)")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("jaros³awski", "przemyski", "le¿ajski", _
            "przeworski", "³añcucki", "ni¿añski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-sêdziszowski", "mielecki", "dêbicki", _
            "brzozowski", "kroœnieñski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy¿owski", "lubaczowski")).Fill
            
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 64, 64)
        End With
        
        With ActiveSheet.Shapes.Range(Array("jaros³awski", "przemyski", "le¿ajski", _
            "przeworski", "³añcucki", "ni¿añski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-sêdziszowski", "mielecki", "dêbicki", _
            "brzozowski", "kroœnieñski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy¿owski", "lubaczowski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("d¹browski", "gorlicki", "nowos¹decki", _
            "bocheñski", "brzeski (ma³opolskie)", "tarnowski", "limanowski", "wielicki", _
            "myœlenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "oœwiêcimski", "chrzanowski", "okulski", "nowotarski", "tatrzañski")).Fill
            
            .ForeColor.RGB = RGB(204, 51, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("d¹browski", "gorlicki", "nowos¹decki", _
            "bocheñski", "brzeski (ma³opolskie)", "tarnowski", "limanowski", "wielicki", _
            "myœlenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "oœwiêcimski", "chrzanowski", "okulski", "nowotarski", "tatrzañski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("Gdynia", "Sopot", "Elbl¹g", "S³upsk", _
            "Koszalin", "Grudzi¹dz", "Œwinoujœcie", "Suwa³ki", "£om¿a", "Ostro³êka", "P³ock", "W³oc³awek" _
            , "Konin", "Siedlce", "Bia³a Podlaska", "Skierniewice", "Piotrków Trybunalski" _
            , "Kalisz", "Leszno", "Legnica", "Jelenia Góra", "Radom", "Che³m", "Zamoœæ", _
            "Tarnobrzeg", "Tarnów", "Nowy S¹cz", "Krosno", "Przemyœl", "Sosnowiec", _
            "D¹browa-Górnicza", "Jaworzno", "Bielsko-Bia³a", "Mys³owice", "Tychy", "¯ory", _
            "Jastrzêbie-Zdrój", "Rybnik", "Ruda Œl¹ska", "Bytom", "Piekary Œl¹skie", _
            "Zabrze", "Gliwice", "Czêstochowa", "Œwiêtoch³owice")).Fill
            
            .ForeColor.RGB = RGB(255, 69, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Gdynia", "Sopot", "Elbl¹g", "S³upsk", _
            "Koszalin", "Œwinoujœcie", "Grudzi¹dz", "Suwa³ki", "£om¿a", "Ostro³êka", "P³ock", "W³oc³awek" _
            , "Konin", "Siedlce", "Bia³a Podlaska", "Skierniewice", "Piotrków Trybunalski" _
            , "Kalisz", "Leszno", "Legnica", "Jelenia Góra", "Radom", "Che³m", "Zamoœæ", _
            "Tarnobrzeg", "Tarnów", "Nowy S¹cz", "Krosno", "Przemyœl", "Sosnowiec", _
            "D¹browa-Górnicza", "Jaworzno", "Bielsko-Bia³a", "Mys³owice", "Tychy", "¯ory", _
            "Jastrzêbie-Zdrój", "Rybnik", "Ruda Œl¹ska", "Bytom", "Piekary Œl¹skie", _
            "Zabrze", "Gliwice", "Czêstochowa", "Œwiêtoch³owice")).Line
            
            .ForeColor.RGB = RGB(0, 0, 0)
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("Gdañsk", "Szczecin", "Gorzów Wielkopolski", _
            "Poznañ", "Bydgoszcz", "Toruñ", "Olsztyn", "Bia³ystok", "Warszawa", "£ódŸ", _
            "Zielona Góra", "Wroc³aw", "Opole", "Kielce", "Lublin", "Rzeszów", "Katowice", _
            "Kraków")).Fill
            
            .ForeColor.RGB = RGB(255, 102, 153)
         End With
         
        With ActiveSheet.Shapes.Range(Array("Gdañsk", "Szczecin", "Gorzów Wielkopolski", _
            "Poznañ", "Bydgoszcz", "Toruñ", "Olsztyn", "Bia³ystok", "Warszawa", "£ódŸ", _
            "Zielona Góra", "Wroc³aw", "Opole", "Kielce", "Lublin", "Rzeszów", "Katowice", _
            "Kraków")).Line
            
            .Weight = 1
            .ForeColor.ObjectThemeColor = msoThemeColorText1
    
        End With
       
        'Wykres 2011
        With ActiveSheet.ChartObjects("Wykres 399")

            With .Chart.SeriesCollection(1)

               .Values = Worksheets("Mapka dane").Range("C" & 383 & ":E" & 383)
        
            End With
    
        End With
    
        'Wykres 2035
        With ActiveSheet.ChartObjects("Wykres 784")

            With .Chart.SeriesCollection(1)

                .Values = Worksheets("Mapka dane").Range("BW" & 383 & ":BY" & 383)
        
            End With
        
        End With
    
        'Wykres dynamiki
        With ActiveSheet.ChartObjects("Wykres 1").Chart
    
            .ChartTitle.Text = "Miasto na prawach powiatu Warszawa - prognoza liczby ludnoœci na lata 2011-2035"
            .SeriesCollection(1).Values = Worksheets("Mapka dane").Range("C" & 383 & "," & "F" & 383 & "," & "I" & 383 & "," & "L" & 383 & "," & "O" & 383 & "," & "R" & 383 & "," & "U" & 383 & "," & "X" & 383 & "," & "AA" & 383 & "," & "AD" & 383 & "," & "AG" & 383 & "," & "AJ" & 383 & "," & "AM" & 383 & "," & "AP" & 383 & "," & "AS" & 383 & "," & "AV" & 383 & "," & "AY" & 383 & "," & "BB" & 383 & "," & "BE" & 383 & "," & "BK" & 383 & "," & "BN" & 383 & "," & "BQ" & 383 & "," & "BT" & 383 & "," & "BW" & 383)
            .SeriesCollection(2).Values = Worksheets("Mapka dane").Range("D" & 383 & "," & "G" & 383 & "," & "J" & 383 & "," & "M" & 383 & "," & "P" & 383 & "," & "S" & 383 & "," & "V" & 383 & "," & "Y" & 383 & "," & "AB" & 383 & "," & "AE" & 383 & "," & "AH" & 383 & "," & "AK" & 383 & "," & "AN" & 383 & "," & "AQ" & 383 & "," & "AT" & 383 & "," & "AW" & 383 & "," & "AZ" & 383 & "," & "BC" & 383 & "," & "BF" & 383 & "," & "BL" & 383 & "," & "BO" & 383 & "," & "BR" & 383 & "," & "BU" & 383 & "," & "BX" & 383)
            .SeriesCollection(3).Values = Worksheets("Mapka dane").Range("E" & 383 & "," & "H" & 383 & "," & "K" & 383 & "," & "N" & 383 & "," & "Q" & 383 & "," & "T" & 383 & "," & "W" & 383 & "," & "Z" & 383 & "," & "AC" & 383 & "," & "AF" & 383 & "," & "AI" & 383 & "," & "AL" & 383 & "," & "AO" & 383 & "," & "AR" & 383 & "," & "AU" & 383 & "," & "AX" & 383 & "," & "BA" & 383 & "," & "BD" & 383 & "," & "BG" & 383 & "," & "BM" & 383 & "," & "BP" & 383 & "," & "BS" & 383 & "," & "BV" & 383 & "," & "BY" & 383)
    
        End With
    
    ElseIf ActiveSheet.Shapes.Range(Array("Woj")).Glow.Radius Then
    
        With ActiveSheet.Shapes("pole tekstowe 2")
    
            .Visible = False
            .TextFrame.Characters.Text = ""
    
        End With
    
        counter = ""
        
        With Sheets("Powiaty").Shapes(counter)

            .Fill.Transparency = 0

        End With
    
        ActiveSheet.Shapes.Range(Array("pomorskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("zachodniopomorskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("warmiñsko-mazurskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("kujawsko-pomorskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("wielkopolskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("lubuskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("dolnoœl¹skie")).Ungroup
        ActiveSheet.Shapes.Range(Array("opolskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("œl¹skie")).Ungroup
        ActiveSheet.Shapes.Range(Array("mazowieckie")).Ungroup
        ActiveSheet.Shapes.Range(Array("podlaskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("œwiêtokrzyskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("lubelskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("³ódzkie")).Ungroup
        ActiveSheet.Shapes.Range(Array("podkarpackie")).Ungroup
        ActiveSheet.Shapes.Range(Array("ma³opolskie")).Ungroup
    
        ActiveSheet.Shapes.Range(Array("Kraj")).Glow.Radius = 0
        ActiveSheet.Shapes.Range(Array("Woj")).Glow.Radius = 0
   
        With ActiveSheet.Shapes.Range(Array("Powiat")).Glow
            .color.ObjectThemeColor = msoThemeColorBackground1
            ' .Transparency = 0.2
            .Radius = 5
        End With
        
        Range("H42:K43").FormulaR1C1 = "GOP"
        Range("A44:S55").Borders(xlEdgeLeft).LineStyle = xlContinuous
        Range("A44:S55").Borders(xlEdgeTop).LineStyle = xlContinuous
        Range("A44:S55").Borders(xlEdgeRight).LineStyle = xlContinuous
           
        ActiveSheet.Shapes.Range(Array(" ")).Visible = msoTrue
           
        ActiveSheet.Shapes.Range(Array(" ")).Ungroup
       
        With ActiveSheet.Shapes.Range(Array("pucki", "wejherowski", "lêborski", "s³upski" _
            , "bytowski", "cz³uchowski", "chojnicki", "koœcierski", "kartuski", "gdañski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzyñski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("s³awieñski", "bia³ogardzki", "koszaliñski", _
            "szczecinecki", "wa³ecki", "drawski", "œwidwiñski", "ko³obrzeski", "gryficki", _
            "³obeski", "Œwinoujœcie", "goleniowski", "stargardzki", "choszczeñski", _
            "pyrzycki", "myœliborski", "gryfiñski", "policki", "kamieñski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("braniewski", "elbl¹ski", "i³awski", _
            "nowomiejski", "dzia³dowski", "ostródzki", "lidzbarski", "bartoszycki", _
            "olsztyñski", "nidzicki", "szczycieñski", "mr¹gowski", "kêtrzyñski", _
            "wêgorzewski", "go³dapski", "olecki", "gi¿ycki", "e³cki", "piski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("sêpoleñski", "tucholski", "che³miñski", _
            "œwiecki", "grudzi¹dzki", "w¹brzeski", "brodnicki", "rypiñski", _
            "golubsko-dobrzyñski", "lipnowski", "w³oc³awski", "radziejowski", _
            "aleksandrowski", "toruñski", "inowroc³awski", "mogileñski", "bydgoski", _
            "¿niñski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("z³otowski", "pilski", "nakielski", _
            "chodzieski", "w¹growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "miêdzychodzki", "nowotomyski", "wolsztyñski", "grodziski (wielkopolskie)", _
            "poznañski", "gnieŸnieñski", "kolski", "koniñski", "s³upecki", "wrzesiñski", _
            "œremski", "jarociñski", "œredzki (wielkopolskie)", "leszczyñski", "gostyñski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszyñski", "ostrzeszowski", "kêpiñski", "turecki", "koœciañski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
     
        With ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "gorzowski", _
            "sulêciñski", "s³ubicki", "miêdzyrzecki", "œwiebodziñski", "zielonogórski", _
            "kroœnieñski (lubuskie)", "¿arski", "nowosolski", "¿agañski", "wschowski")). _
            Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("wa³brzyski", "kamiennogórski", "z³otoryjski" _
            , "lubañski", "zgorzelecki", "boles³awiecki", "polkowicki", "g³ogowski", _
            "górowski", "wo³owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "oleœnicki", "wroc³awski", "o³awski", "strzeliñski", "dzier¿oniowski", _
            "z¹bkowicki", "lubiñski", "œredzki (dolnyœl¹sk)", "œwidnicki (dolnoœl¹skie)", _
            "lwówecki", "jeleniogórski", "k³odzki")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("suwalski", "sejneñski", "augustowski", _
            "grajewski", "kolneñski", "moniecki", "sokólski", "bia³ostocki", "³om¿yñski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("ostro³êcki", "przasnyski", "ciechanowski", _
            "m³awski", "¿uromiñski", "sierpecki", "gostyniñski", "p³ocki", "p³oñski", _
            "pu³tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko³owski", "³osicki", "siedlecki", "wêgrowski", _
            "wo³omiñski", "miñski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "¿yrardowski", "piaseczyñski", "otwocki", "grójecki", "przysuski", _
            "bia³obrzeski", "kozienicki", "garwoliñski", "zwoleñski", "szyd³owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("rawski", "skierniewicki", "³owicki", _
            "kutnowski", "³êczycki", "zgierski", "brzeziñski", "tomaszowski (³ódzkie)", _
            "opoczyñski", "³ódzki wschodni", "piotrkowski", "radomszczañski", "pajêczañski" _
            , "be³chatowski", "pabianicki", "³aski", "zduñskowolski", "poddêbicki", _
            "sieradzki", "wieluñski", "wieruszowski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("g³ubczycki", "kêdzierzyñsko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys³owski", "brzeski (opolskie)", "nyski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("bieruñsko-lêdziñski", "pszczyñski", _
            "bielski (œl¹skie)", "cieszyñski", "wodzis³awski", "rybnicki", "raciborski", _
            "gliwicki", "miko³owski", "bêdziñski", "tarnogórski", "zawierciañski", _
            "myszkowski", "czêstochowski", "lubliniecki", "k³obucki", "¿ywiecki")).Line
            
             .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("konecki", "skar¿yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w³oszczowski", "jêdrzejowski", "piñczowski", "buski", "kazimierski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("bialski", "radzyñski", "³ukowski", "rycki", _
            "parczewski", "w³odawski", "lubartowski", "pu³awski", "opolski (lubelskie)", _
            "lubelski", "kraœnicki", "œwidnicki (lubelskie)", "che³mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi³gorajski", "³êczyñski", _
            "tomaszowski (lubelskie)")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("jaros³awski", "przemyski", "le¿ajski", _
            "przeworski", "³añcucki", "ni¿añski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-sêdziszowski", "mielecki", "dêbicki", _
            "brzozowski", "kroœnieñski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy¿owski", "lubaczowski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("d¹browski", "gorlicki", "nowos¹decki", _
            "bocheñski", "brzeski (ma³opolskie)", "tarnowski", "limanowski", "wielicki", _
            "myœlenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "oœwiêcimski", "chrzanowski", "okulski", "nowotarski", "tatrzañski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("Gdynia", "Sopot", "Elbl¹g", "S³upsk", _
            "Koszalin", "Œwinoujœcie", "Grudzi¹dz", "Suwa³ki", "£om¿a", "Ostro³êka", "P³ock", "W³oc³awek" _
            , "Konin", "Siedlce", "Bia³a Podlaska", "Skierniewice", "Piotrków Trybunalski" _
            , "Kalisz", "Leszno", "Legnica", "Jelenia Góra", "Radom", "Che³m", "Zamoœæ", _
            "Tarnobrzeg", "Tarnów", "Nowy S¹cz", "Krosno", "Przemyœl", "Sosnowiec", _
            "D¹browa-Górnicza", "Jaworzno", "Bielsko-Bia³a", "Mys³owice", "Tychy", "¯ory", _
            "Jastrzêbie-Zdrój", "Rybnik", "Ruda Œl¹ska", "Bytom", "Piekary Œl¹skie", _
            "Zabrze", "Gliwice", "Czêstochowa", "Œwiêtoch³owice")).Fill
            
            .ForeColor.RGB = RGB(255, 69, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Gdynia", "Sopot", "Elbl¹g", "S³upsk", _
            "Koszalin", "Œwinoujœcie", "Grudzi¹dz", "Suwa³ki", "£om¿a", "Ostro³êka", "P³ock", "W³oc³awek" _
            , "Konin", "Siedlce", "Bia³a Podlaska", "Skierniewice", "Piotrków Trybunalski" _
            , "Kalisz", "Leszno", "Legnica", "Jelenia Góra", "Radom", "Che³m", "Zamoœæ", _
            "Tarnobrzeg", "Tarnów", "Nowy S¹cz", "Krosno", "Przemyœl", "Sosnowiec", _
            "D¹browa-Górnicza", "Jaworzno", "Bielsko-Bia³a", "Mys³owice", "Tychy", "¯ory", _
            "Jastrzêbie-Zdrój", "Rybnik", "Ruda Œl¹ska", "Bytom", "Piekary Œl¹skie", _
            "Zabrze", "Gliwice", "Czêstochowa", "Œwiêtoch³owice")).Line
            
            .ForeColor.RGB = RGB(0, 0, 0)
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("Gdañsk", "Szczecin", "Gorzów Wielkopolski", _
            "Poznañ", "Bydgoszcz", "Toruñ", "Olsztyn", "Bia³ystok", "Warszawa", "£ódŸ", _
            "Zielona Góra", "Wroc³aw", "Opole", "Kielce", "Lublin", "Rzeszów", "Katowice", _
            "Kraków")).Fill
            
            .ForeColor.RGB = RGB(255, 102, 153)
         End With
         
        With ActiveSheet.Shapes.Range(Array("Gdañsk", "Szczecin", "Gorzów Wielkopolski", _
            "Poznañ", "Bydgoszcz", "Toruñ", "Olsztyn", "Bia³ystok", "Warszawa", "£ódŸ", _
            "Zielona Góra", "Wroc³aw", "Opole", "Kielce", "Lublin", "Rzeszów", "Katowice", _
            "Kraków")).Line
            
            .Weight = 1
            .ForeColor.ObjectThemeColor = msoThemeColorText1
    
        End With
        
        'Wykres 2011
        With ActiveSheet.ChartObjects("Wykres 399")

            With .Chart.SeriesCollection(1)

               .Values = Worksheets("Mapka dane").Range("C" & 383 & ":E" & 383)
        
            End With
    
        End With
    
        'Wykres 2035
        With ActiveSheet.ChartObjects("Wykres 784")

            With .Chart.SeriesCollection(1)

                .Values = Worksheets("Mapka dane").Range("BW" & 383 & ":BY" & 383)
        
            End With
        
        End With
    
        'Wykres dynamiki
        With ActiveSheet.ChartObjects("Wykres 1").Chart
    
            .ChartTitle.Text = "Miasto na prawach powiatu Warszawa - prognoza liczby ludnoœci na lata 2011-2035"
            .SeriesCollection(1).Values = Worksheets("Mapka dane").Range("C" & 383 & "," & "F" & 383 & "," & "I" & 383 & "," & "L" & 383 & "," & "O" & 383 & "," & "R" & 383 & "," & "U" & 383 & "," & "X" & 383 & "," & "AA" & 383 & "," & "AD" & 383 & "," & "AG" & 383 & "," & "AJ" & 383 & "," & "AM" & 383 & "," & "AP" & 383 & "," & "AS" & 383 & "," & "AV" & 383 & "," & "AY" & 383 & "," & "BB" & 383 & "," & "BE" & 383 & "," & "BK" & 383 & "," & "BN" & 383 & "," & "BQ" & 383 & "," & "BT" & 383 & "," & "BW" & 383)
            .SeriesCollection(2).Values = Worksheets("Mapka dane").Range("D" & 383 & "," & "G" & 383 & "," & "J" & 383 & "," & "M" & 383 & "," & "P" & 383 & "," & "S" & 383 & "," & "V" & 383 & "," & "Y" & 383 & "," & "AB" & 383 & "," & "AE" & 383 & "," & "AH" & 383 & "," & "AK" & 383 & "," & "AN" & 383 & "," & "AQ" & 383 & "," & "AT" & 383 & "," & "AW" & 383 & "," & "AZ" & 383 & "," & "BC" & 383 & "," & "BF" & 383 & "," & "BL" & 383 & "," & "BO" & 383 & "," & "BR" & 383 & "," & "BU" & 383 & "," & "BX" & 383)
            .SeriesCollection(3).Values = Worksheets("Mapka dane").Range("E" & 383 & "," & "H" & 383 & "," & "K" & 383 & "," & "N" & 383 & "," & "Q" & 383 & "," & "T" & 383 & "," & "W" & 383 & "," & "Z" & 383 & "," & "AC" & 383 & "," & "AF" & 383 & "," & "AI" & 383 & "," & "AL" & 383 & "," & "AO" & 383 & "," & "AR" & 383 & "," & "AU" & 383 & "," & "AX" & 383 & "," & "BA" & 383 & "," & "BD" & 383 & "," & "BG" & 383 & "," & "BM" & 383 & "," & "BP" & 383 & "," & "BS" & 383 & "," & "BV" & 383 & "," & "BY" & 383)
    
        End With
          
    End If
    
    RefreshCharts  'odswieza wykresy, ¿eby update'owa³y siê w czasie dzia³ania excela
     
End Sub

Sub CaleWoj()

    

    If ActiveSheet.Shapes.Range(Array("Woj")).Glow.Radius <> 0 Then Exit Sub
    
    If ActiveSheet.Shapes.Range(Array("Kraj")).Glow.Radius <> 0 Then
    
        ActiveSheet.Shapes.Range(Array("Polska")).Ungroup
        
        ActiveSheet.Shapes.Range(Array("Grupa 1")).Ungroup
        ActiveSheet.Shapes.Range(Array("Grupa 2")).Ungroup
        ActiveSheet.Shapes.Range(Array("Grupa 3")).Ungroup
    
        ActiveSheet.Shapes.Range(Array("Kraj")).Glow.Radius = 0
        ActiveSheet.Shapes.Range(Array("Powiat")).Glow.Radius = 0
    
        With ActiveSheet.Shapes.Range(Array("Woj")).Glow
        
            .color.ObjectThemeColor = msoThemeColorBackground1
            ' .Transparency = 0.2
            .Radius = 5
        End With
              
        ActiveSheet.Shapes.Range(Array("pucki", "Gdynia", "Sopot", "Gdañsk", "S³upsk", "wejherowski", "lêborski", "s³upski" _
            , "bytowski", "cz³uchowski", "chojnicki", "koœcierski", "kartuski", "gdañski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzyñski")).Group.Name = "pomorskie"
            
        With ActiveSheet.Shapes.Range(Array("pucki", "Gdynia", "Sopot", "Gdañsk", "S³upsk", "wejherowski", "lêborski", "s³upski" _
            , "bytowski", "cz³uchowski", "chojnicki", "koœcierski", "kartuski", "gdañski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzyñski")).Fill
            
            .ForeColor.RGB = RGB(255, 192, 0)
            .Solid
        End With
        
        With ActiveSheet.Shapes.Range(Array("pucki", "Gdynia", "Sopot", "Gdañsk", "S³upsk", "wejherowski", "lêborski", "s³upski" _
            , "bytowski", "cz³uchowski", "chojnicki", "koœcierski", "kartuski", "gdañski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzyñski")).Line
            
            .ForeColor.RGB = RGB(255, 192, 0)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("s³awieñski", "Szczecin", "Koszalin", "bia³ogardzki", "koszaliñski", _
            "szczecinecki", "wa³ecki", "drawski", "œwidwiñski", "ko³obrzeski", "gryficki", _
            "³obeski", "Œwinoujœcie", "goleniowski", "stargardzki", "choszczeñski", _
            "pyrzycki", "myœliborski", "gryfiñski", "policki", "kamieñski")).Group.Name = "zachodniopomorskie"
    
        With ActiveSheet.Shapes.Range(Array("s³awieñski", "Szczecin", "Koszalin", "bia³ogardzki", "koszaliñski", _
            "szczecinecki", "wa³ecki", "drawski", "œwidwiñski", "ko³obrzeski", "gryficki", _
            "³obeski", "Œwinoujœcie", "goleniowski", "stargardzki", "choszczeñski", _
            "pyrzycki", "myœliborski", "gryfiñski", "policki", "kamieñski")).Fill
            
            .ForeColor.RGB = RGB(192, 0, 0)
        End With

        With ActiveSheet.Shapes.Range(Array("s³awieñski", "Szczecin", "Koszalin", "bia³ogardzki", "koszaliñski", _
            "szczecinecki", "wa³ecki", "drawski", "œwidwiñski", "ko³obrzeski", "gryficki", _
            "³obeski", "Œwinoujœcie", "goleniowski", "stargardzki", "choszczeñski", _
            "pyrzycki", "myœliborski", "gryfiñski", "policki", "kamieñski")).Line
            
            .ForeColor.RGB = RGB(192, 0, 0)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Elbl¹g", "Olsztyn", "braniewski", "elbl¹ski", "i³awski", _
            "nowomiejski", "dzia³dowski", "ostródzki", "lidzbarski", "bartoszycki", _
            "olsztyñski", "nidzicki", "szczycieñski", "mr¹gowski", "kêtrzyñski", _
            "wêgorzewski", "go³dapski", "olecki", "gi¿ycki", "e³cki", "piski")).Group.Name = "warmiñsko-mazurskie"
    
        With ActiveSheet.Shapes.Range(Array("Elbl¹g", "Olsztyn", "braniewski", "elbl¹ski", "i³awski", _
            "nowomiejski", "dzia³dowski", "ostródzki", "lidzbarski", "bartoszycki", _
            "olsztyñski", "nidzicki", "szczycieñski", "mr¹gowski", "kêtrzyñski", _
            "wêgorzewski", "go³dapski", "olecki", "gi¿ycki", "e³cki", "piski")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent3
        End With
        
        With ActiveSheet.Shapes.Range(Array("Elbl¹g", "Olsztyn", "braniewski", "elbl¹ski", "i³awski", _
            "nowomiejski", "dzia³dowski", "ostródzki", "lidzbarski", "bartoszycki", _
            "olsztyñski", "nidzicki", "szczycieñski", "mr¹gowski", "kêtrzyñski", _
            "wêgorzewski", "go³dapski", "olecki", "gi¿ycki", "e³cki", "piski")).Line
            .ForeColor.ObjectThemeColor = msoThemeColorAccent3
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("sêpoleñski", "tucholski", "Bydgoszcz", "Toruñ", "Grudzi¹dz", "W³oc³awek", "che³miñski", _
            "œwiecki", "grudzi¹dzki", "w¹brzeski", "brodnicki", "rypiñski", _
            "golubsko-dobrzyñski", "lipnowski", "w³oc³awski", "radziejowski", _
            "aleksandrowski", "toruñski", "inowroc³awski", "mogileñski", "bydgoski", _
            "¿niñski")).Group.Name = "kujawsko-pomorskie"
            
        With ActiveSheet.Shapes.Range(Array("sêpoleñski", "tucholski", "Bydgoszcz", "Toruñ", "Grudzi¹dz", "W³oc³awek", "che³miñski", _
            "œwiecki", "grudzi¹dzki", "w¹brzeski", "brodnicki", "rypiñski", _
            "golubsko-dobrzyñski", "lipnowski", "w³oc³awski", "radziejowski", _
            "aleksandrowski", "toruñski", "inowroc³awski", "mogileñski", "bydgoski", _
            "¿niñski")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent4
        End With
        
        With ActiveSheet.Shapes.Range(Array("sêpoleñski", "tucholski", "Bydgoszcz", "Toruñ", "Grudzi¹dz", "W³oc³awek", "che³miñski", _
            "œwiecki", "grudzi¹dzki", "w¹brzeski", "brodnicki", "rypiñski", _
            "golubsko-dobrzyñski", "lipnowski", "w³oc³awski", "radziejowski", _
            "aleksandrowski", "toruñski", "inowroc³awski", "mogileñski", "bydgoski", _
            "¿niñski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent4
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Poznañ", "Konin", "Kalisz", "Leszno", "z³otowski", "pilski", "nakielski", _
            "chodzieski", "w¹growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "miêdzychodzki", "nowotomyski", "wolsztyñski", "grodziski (wielkopolskie)", _
            "poznañski", "gnieŸnieñski", "kolski", "koniñski", "s³upecki", "wrzesiñski", _
            "œremski", "jarociñski", "œredzki (wielkopolskie)", "leszczyñski", "gostyñski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszyñski", "ostrzeszowski", "kêpiñski", "turecki", "koœciañski")).Group.Name = "wielkopolskie"
    
        With ActiveSheet.Shapes.Range(Array("Poznañ", "Konin", "Kalisz", "Leszno", "z³otowski", "pilski", "nakielski", _
            "chodzieski", "w¹growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "miêdzychodzki", "nowotomyski", "wolsztyñski", "grodziski (wielkopolskie)", _
            "poznañski", "gnieŸnieñski", "kolski", "koniñski", "s³upecki", "wrzesiñski", _
            "œremski", "jarociñski", "œredzki (wielkopolskie)", "leszczyñski", "gostyñski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszyñski", "ostrzeszowski", "kêpiñski", "turecki", "koœciañski")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent5
        End With
        
        With ActiveSheet.Shapes.Range(Array("Poznañ", "Konin", "Kalisz", "Leszno", "z³otowski", "pilski", "nakielski", _
            "chodzieski", "w¹growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "miêdzychodzki", "nowotomyski", "wolsztyñski", "grodziski (wielkopolskie)", _
            "poznañski", "gnieŸnieñski", "kolski", "koniñski", "s³upecki", "wrzesiñski", _
            "œremski", "jarociñski", "œredzki (wielkopolskie)", "leszczyñski", "gostyñski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszyñski", "ostrzeszowski", "kêpiñski", "turecki", "koœciañski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent5
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "Gorzów Wielkopolski", "Zielona Góra", "gorzowski", _
            "sulêciñski", "s³ubicki", "miêdzyrzecki", "œwiebodziñski", "zielonogórski", _
            "kroœnieñski (lubuskie)", "¿arski", "nowosolski", "¿agañski", "wschowski")). _
            Group.Name = "lubuskie"
        
        With ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "Gorzów Wielkopolski", "Zielona Góra", "gorzowski", _
            "sulêciñski", "s³ubicki", "miêdzyrzecki", "œwiebodziñski", "zielonogórski", _
            "kroœnieñski (lubuskie)", "¿arski", "nowosolski", "¿agañski", "wschowski")).Fill
            
            .ForeColor.RGB = RGB(165, 27, 96)
        End With
        
        With ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "Gorzów Wielkopolski", "Zielona Góra", "gorzowski", _
            "sulêciñski", "s³ubicki", "miêdzyrzecki", "œwiebodziñski", "zielonogórski", _
            "kroœnieñski (lubuskie)", "¿arski", "nowosolski", "¿agañski", "wschowski")).Line
            
            .ForeColor.RGB = RGB(165, 27, 96)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Wroc³aw", "Jelenia Góra", "Legnica", "wa³brzyski", "kamiennogórski", "z³otoryjski" _
            , "lubañski", "zgorzelecki", "boles³awiecki", "polkowicki", "g³ogowski", _
            "górowski", "wo³owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "oleœnicki", "wroc³awski", "o³awski", "strzeliñski", "dzier¿oniowski", _
            "z¹bkowicki", "lubiñski", "œredzki (dolnyœl¹sk)", "œwidnicki (dolnoœl¹skie)", _
            "lwówecki", "jeleniogórski", "k³odzki")).Group.Name = "dolnoœl¹skie"
    
        With ActiveSheet.Shapes.Range(Array("Wroc³aw", "Jelenia Góra", "Legnica", "wa³brzyski", "kamiennogórski", "z³otoryjski" _
            , "lubañski", "zgorzelecki", "boles³awiecki", "polkowicki", "g³ogowski", _
            "górowski", "wo³owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "oleœnicki", "wroc³awski", "o³awski", "strzeliñski", "dzier¿oniowski", _
            "z¹bkowicki", "lubiñski", "œredzki (dolnyœl¹sk)", "œwidnicki (dolnoœl¹skie)", _
            "lwówecki", "jeleniogórski", "k³odzki")).Fill
            
            .ForeColor.RGB = RGB(255, 0, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Wroc³aw", "Jelenia Góra", "Legnica", "wa³brzyski", "kamiennogórski", "z³otoryjski" _
            , "lubañski", "zgorzelecki", "boles³awiecki", "polkowicki", "g³ogowski", _
            "górowski", "wo³owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "oleœnicki", "wroc³awski", "o³awski", "strzeliñski", "dzier¿oniowski", _
            "z¹bkowicki", "lubiñski", "œredzki (dolnyœl¹sk)", "œwidnicki (dolnoœl¹skie)", _
            "lwówecki", "jeleniogórski", "k³odzki")).Line
            
            .ForeColor.RGB = RGB(255, 0, 0)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Suwa³ki", "Bia³ystok", "£om¿a", "suwalski", "sejneñski", "augustowski", _
            "grajewski", "kolneñski", "moniecki", "sokólski", "bia³ostocki", "³om¿yñski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Group.Name = "podlaskie"
        
        With ActiveSheet.Shapes.Range(Array("Suwa³ki", "Bia³ystok", "£om¿a", "suwalski", "sejneñski", "augustowski", _
            "grajewski", "kolneñski", "moniecki", "sokólski", "bia³ostocki", "³om¿yñski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        End With
        
        With ActiveSheet.Shapes.Range(Array("Suwa³ki", "Bia³ystok", "£om¿a", "suwalski", "sejneñski", "augustowski", _
            "grajewski", "kolneñski", "moniecki", "sokólski", "bia³ostocki", "³om¿yñski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Ostro³êka", "P³ock", "Warszawa", "Radom", "Siedlce", "ostro³êcki", "przasnyski", "ciechanowski", _
            "m³awski", "¿uromiñski", "sierpecki", "gostyniñski", "p³ocki", "p³oñski", _
            "pu³tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko³owski", "³osicki", "siedlecki", "wêgrowski", _
            "wo³omiñski", "miñski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "¿yrardowski", "piaseczyñski", "otwocki", "grójecki", "przysuski", _
            "bia³obrzeski", "kozienicki", "garwoliñski", "zwoleñski", "szyd³owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Group.Name = "mazowieckie"
            
        With ActiveSheet.Shapes.Range(Array("Ostro³êka", "P³ock", "Warszawa", "Radom", "Siedlce", "ostro³êcki", "przasnyski", "ciechanowski", _
            "m³awski", "¿uromiñski", "sierpecki", "gostyniñski", "p³ocki", "p³oñski", _
            "pu³tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko³owski", "³osicki", "siedlecki", "wêgrowski", _
            "wo³omiñski", "miñski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "¿yrardowski", "piaseczyñski", "otwocki", "grójecki", "przysuski", _
            "bia³obrzeski", "kozienicki", "garwoliñski", "zwoleñski", "szyd³owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Fill
            
            .ForeColor.RGB = RGB(255, 204, 51)
            
         End With
         
        With ActiveSheet.Shapes.Range(Array("Ostro³êka", "P³ock", "Warszawa", "Radom", "Siedlce", "ostro³êcki", "przasnyski", "ciechanowski", _
            "m³awski", "¿uromiñski", "sierpecki", "gostyniñski", "p³ocki", "p³oñski", _
            "pu³tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko³owski", "³osicki", "siedlecki", "wêgrowski", _
            "wo³omiñski", "miñski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "¿yrardowski", "piaseczyñski", "otwocki", "grójecki", "przysuski", _
            "bia³obrzeski", "kozienicki", "garwoliñski", "zwoleñski", "szyd³owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Line
            
            .ForeColor.RGB = RGB(255, 204, 51)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("rawski", "£ódŸ", "Piotrków Trybunalski", "Skierniewice", "skierniewicki", "³owicki", _
            "kutnowski", "³êczycki", "zgierski", "brzeziñski", "tomaszowski (³ódzkie)", _
            "opoczyñski", "³ódzki wschodni", "piotrkowski", "radomszczañski", "pajêczañski" _
            , "be³chatowski", "pabianicki", "³aski", "zduñskowolski", "poddêbicki", _
            "sieradzki", "wieluñski", "wieruszowski")).Group.Name = "³ódzkie"
        
        With ActiveSheet.Shapes.Range(Array("rawski", "£ódŸ", "Piotrków Trybunalski", "Skierniewice", "skierniewicki", "³owicki", _
            "kutnowski", "³êczycki", "zgierski", "brzeziñski", "tomaszowski (³ódzkie)", _
            "opoczyñski", "³ódzki wschodni", "piotrkowski", "radomszczañski", "pajêczañski" _
            , "be³chatowski", "pabianicki", "³aski", "zduñskowolski", "poddêbicki", _
            "sieradzki", "wieluñski", "wieruszowski")).Fill
            
            .ForeColor.RGB = RGB(0, 176, 80)
        End With
        
        With ActiveSheet.Shapes.Range(Array("rawski", "£ódŸ", "Piotrków Trybunalski", "Skierniewice", "skierniewicki", "³owicki", _
            "kutnowski", "³êczycki", "zgierski", "brzeziñski", "tomaszowski (³ódzkie)", _
            "opoczyñski", "³ódzki wschodni", "piotrkowski", "radomszczañski", "pajêczañski" _
            , "be³chatowski", "pabianicki", "³aski", "zduñskowolski", "poddêbicki", _
            "sieradzki", "wieluñski", "wieruszowski")).Line
            
            .ForeColor.RGB = RGB(0, 176, 80)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Opole", "g³ubczycki", "kêdzierzyñsko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys³owski", "brzeski (opolskie)", "nyski")).Group.Name = "opolskie"
        
        With ActiveSheet.Shapes.Range(Array("Opole", "g³ubczycki", "kêdzierzyñsko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys³owski", "brzeski (opolskie)", "nyski")).Fill
            
            .ForeColor.RGB = RGB(0, 0, 204)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Opole", "g³ubczycki", "kêdzierzyñsko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys³owski", "brzeski (opolskie)", "nyski")).Line
            
            .ForeColor.RGB = RGB(0, 0, 204)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Czêstochowa", "Bielsko-Bia³a", "Ruda Œl¹ska", "Œwiêtoch³owice", "Chorzów" _
        , "Siemianowice Œl¹skie", "Piekary Œl¹skie", "Bytom", "Zabrze", "Gliwice", "Rybnik", "Jastrzêbie-Zdrój", "¯ory", _
        "Tychy", "Jaworzno", "Mys³owice", "Katowice", "Sosnowiec", "D¹browa-Górnicza", "bieruñsko-lêdziñski", "pszczyñski", _
        "bielski (œl¹skie)", "cieszyñski", "wodzis³awski", "rybnicki", "raciborski", _
        "gliwicki", "miko³owski", "bêdziñski", "tarnogórski", "zawierciañski", _
        "myszkowski", "czêstochowski", "lubliniecki", "k³obucki", "¿ywiecki")).Group.Name = "œl¹skie"
    
        With ActiveSheet.Shapes.Range(Array("Czêstochowa", "Bielsko-Bia³a", "Ruda Œl¹ska", "Œwiêtoch³owice", "Chorzów" _
        , "Siemianowice Œl¹skie", "Piekary Œl¹skie", "Bytom", "Zabrze", "Gliwice", "Rybnik", "Jastrzêbie-Zdrój", "¯ory", _
        "Tychy", "Jaworzno", "Mys³owice", "Katowice", "Sosnowiec", "D¹browa-Górnicza", "bieruñsko-lêdziñski", "pszczyñski", _
        "bielski (œl¹skie)", "cieszyñski", "wodzis³awski", "rybnicki", "raciborski", _
        "gliwicki", "miko³owski", "bêdziñski", "tarnogórski", "zawierciañski", _
        "myszkowski", "czêstochowski", "lubliniecki", "k³obucki", "¿ywiecki")).Fill
        
            .ForeColor.RGB = RGB(204, 153, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Czêstochowa", "Bielsko-Bia³a", "Ruda Œl¹ska", "Œwiêtoch³owice", "Chorzów" _
        , "Siemianowice Œl¹skie", "Piekary Œl¹skie", "Bytom", "Zabrze", "Gliwice", "Rybnik", "Jastrzêbie-Zdrój", "¯ory", _
        "Tychy", "Jaworzno", "Mys³owice", "Katowice", "Sosnowiec", "D¹browa-Górnicza", "bieruñsko-lêdziñski", "pszczyñski", _
        "bielski (œl¹skie)", "cieszyñski", "wodzis³awski", "rybnicki", "raciborski", _
        "gliwicki", "miko³owski", "bêdziñski", "tarnogórski", "zawierciañski", _
        "myszkowski", "czêstochowski", "lubliniecki", "k³obucki", "¿ywiecki")).Line
        
             .ForeColor.RGB = RGB(204, 153, 0)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Kielce", "konecki", "skar¿yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w³oszczowski", "jêdrzejowski", "piñczowski", "buski", "kazimierski")).Group.Name = "œwiêtokrzyskie"
            
        With ActiveSheet.Shapes.Range(Array("Kielce", "konecki", "skar¿yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w³oszczowski", "jêdrzejowski", "piñczowski", "buski", "kazimierski")).Fill
            
            .ForeColor.RGB = RGB(204, 51, 255)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Kielce", "konecki", "skar¿yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w³oszczowski", "jêdrzejowski", "piñczowski", "buski", "kazimierski")).Line
            
            .ForeColor.RGB = RGB(204, 51, 255)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Zamoœæ", "Che³m", "Bia³a Podlaska", "Lublin", "bialski", "radzyñski", "³ukowski", "rycki", _
            "parczewski", "w³odawski", "lubartowski", "pu³awski", "opolski (lubelskie)", _
            "lubelski", "kraœnicki", "œwidnicki (lubelskie)", "che³mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi³gorajski", "³êczyñski", _
            "tomaszowski (lubelskie)")).Group.Name = "lubelskie"
            
        With ActiveSheet.Shapes.Range(Array("Zamoœæ", "Che³m", "Bia³a Podlaska", "Lublin", "bialski", "radzyñski", "³ukowski", "rycki", _
            "parczewski", "w³odawski", "lubartowski", "pu³awski", "opolski (lubelskie)", _
            "lubelski", "kraœnicki", "œwidnicki (lubelskie)", "che³mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi³gorajski", "³êczyñski", _
            "tomaszowski (lubelskie)")).Fill
            
            .ForeColor.RGB = RGB(102, 0, 204)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Zamoœæ", "Che³m", "Bia³a Podlaska", "Lublin", "bialski", "radzyñski", "³ukowski", "rycki", _
            "parczewski", "w³odawski", "lubartowski", "pu³awski", "opolski (lubelskie)", _
            "lubelski", "kraœnicki", "œwidnicki (lubelskie)", "che³mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi³gorajski", "³êczyñski", _
            "tomaszowski (lubelskie)")).Line
            
            .ForeColor.RGB = RGB(102, 0, 204)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Rzeszów", "Przemyœl", "Krosno", "Tarnobrzeg", "jaros³awski", "przemyski", "le¿ajski", _
            "przeworski", "³añcucki", "ni¿añski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-sêdziszowski", "mielecki", "dêbicki", _
            "brzozowski", "kroœnieñski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy¿owski", "lubaczowski")).Group.Name = "podkarpackie"
        
        With ActiveSheet.Shapes.Range(Array("Rzeszów", "Przemyœl", "Krosno", "Tarnobrzeg", "jaros³awski", "przemyski", "le¿ajski", _
            "przeworski", "³añcucki", "ni¿añski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-sêdziszowski", "mielecki", "dêbicki", _
            "brzozowski", "kroœnieñski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy¿owski", "lubaczowski")).Fill
            
            .ForeColor.RGB = RGB(255, 64, 64)
            
        End With
        
        With ActiveSheet.Shapes.Range(Array("Rzeszów", "Przemyœl", "Krosno", "Tarnobrzeg", "jaros³awski", "przemyski", "le¿ajski", _
            "przeworski", "³añcucki", "ni¿añski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-sêdziszowski", "mielecki", "dêbicki", _
            "brzozowski", "kroœnieñski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy¿owski", "lubaczowski")).Line
            
            .ForeColor.RGB = RGB(255, 64, 64)
            .Weight = 2.5
        End With

        ActiveSheet.Shapes.Range(Array("Kraków", "Tarnów", "Nowy S¹cz", "d¹browski", "gorlicki", "nowos¹decki", _
            "bocheñski", "brzeski (ma³opolskie)", "tarnowski", "limanowski", "wielicki", _
            "myœlenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "oœwiêcimski", "chrzanowski", "okulski", "nowotarski", "tatrzañski")).Group.Name = "ma³opolskie"
                
        With ActiveSheet.Shapes.Range(Array("Kraków", "Tarnów", "Nowy S¹cz", "d¹browski", "gorlicki", "nowos¹decki", _
            "bocheñski", "brzeski (ma³opolskie)", "tarnowski", "limanowski", "wielicki", _
            "myœlenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "oœwiêcimski", "chrzanowski", "okulski", "nowotarski", "tatrzañski")).Fill
            
            .ForeColor.RGB = RGB(204, 51, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Kraków", "Tarnów", "Nowy S¹cz", "d¹browski", "gorlicki", "nowos¹decki", _
            "bocheñski", "brzeski (ma³opolskie)", "tarnowski", "limanowski", "wielicki", _
            "myœlenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "oœwiêcimski", "chrzanowski", "okulski", "nowotarski", "tatrzañski")).Line
            
            .ForeColor.RGB = RGB(204, 51, 0)
            .Weight = 2.5
        End With
               
        'Wykres 2011
        With ActiveSheet.ChartObjects("Wykres 399")

            With .Chart.SeriesCollection(1)

               .Values = Worksheets("Mapka dane").Range("C" & 426 & ":E" & 426)
            
            End With
    
        End With
    
        'Wykres 2035
        With ActiveSheet.ChartObjects("Wykres 784")
 
            With .Chart.SeriesCollection(1)

                .Values = Worksheets("Mapka dane").Range("BW" & 426 & ":BY" & poz + 426)
        
            End With
        
        End With
        'Wykres dynamiki
        With ActiveSheet.ChartObjects("Wykres 1").Chart
    
            .ChartTitle.Text = "Województwo mazowieckie - prognoza liczby ludnoœci na lata 2011-2035"
            .SeriesCollection(1).Values = Worksheets("Mapka dane").Range("C" & 426 & "," & "F" & 426 & "," & "I" & 426 & "," & "L" & 426 & "," & "O" & 426 & "," & "R" & 426 & "," & "U" & 426 & "," & "X" & 426 & "," & "AA" & 426 & "," & "AD" & 426 & "," & "AG" & 426 & "," & "AJ" & 426 & "," & "AM" & 426 & "," & "AP" & 426 & "," & "AS" & 426 & "," & "AV" & 426 & "," & "AY" & 426 & "," & "BB" & 426 & "," & "BE" & 426 & "," & "BK" & 426 & "," & "BN" & 426 & "," & "BQ" & 426 & "," & "BT" & 426 & "," & "BW" & 426)
            .SeriesCollection(2).Values = Worksheets("Mapka dane").Range("D" & 426 & "," & "G" & 426 & "," & "J" & 426 & "," & "M" & 426 & "," & "P" & 426 & "," & "S" & 426 & "," & "V" & 426 & "," & "Y" & 426 & "," & "AB" & 426 & "," & "AE" & 426 & "," & "AH" & 426 & "," & "AK" & 426 & "," & "AN" & 426 & "," & "AQ" & 426 & "," & "AT" & 426 & "," & "AW" & 426 & "," & "AZ" & 426 & "," & "BC" & 426 & "," & "BF" & 426 & "," & "BL" & 426 & "," & "BO" & 426 & "," & "BR" & 426 & "," & "BU" & 426 & "," & "BX" & 426)
            .SeriesCollection(3).Values = Worksheets("Mapka dane").Range("E" & 426 & "," & "H" & 426 & "," & "K" & 426 & "," & "N" & 426 & "," & "Q" & 426 & "," & "T" & 426 & "," & "W" & 426 & "," & "Z" & 426 & "," & "AC" & 426 & "," & "AF" & 426 & "," & "AI" & 426 & "," & "AL" & 426 & "," & "AO" & 426 & "," & "AR" & 426 & "," & "AU" & 426 & "," & "AX" & 426 & "," & "BA" & 426 & "," & "BD" & 426 & "," & "BG" & 426 & "," & "BM" & 426 & "," & "BP" & 426 & "," & "BS" & 426 & "," & "BV" & 426 & "," & "BY" & 426)
    
        End With
    
    ElseIf ActiveSheet.Shapes.Range(Array("Powiat")).Glow.Radius <> 0 Then
    
        With ActiveSheet.Shapes("pole tekstowe 2")
    
            .Visible = False
            .TextFrame.Characters.Text = ""
    
        End With
        
        counter = ""
        
        With Sheets("Powiaty").Shapes(counter)

            .Fill.Transparency = 0

        End With
                
        ActiveSheet.Shapes.Range(Array("Kraj")).Glow.Radius = 0
        ActiveSheet.Shapes.Range(Array("Powiat")).Glow.Radius = 0
    
        Range("H42:K43").FormulaR1C1 = ""
        Range("A44:S55").Borders(xlEdgeLeft).LineStyle = xlNone
        Range("A44:S55").Borders(xlEdgeTop).LineStyle = xlNone
        Range("A44:S55").Borders(xlEdgeRight).LineStyle = xlNone
        
        ActiveSheet.Shapes.Range(Array("Ruda Œl¹ska ", "Œwiêtoch³owice ", "Chorzów " _
        , "bêdziñski ", "Siemianowice Œl¹skie ", "Piekary Œl¹skie ", "Bytom ", _
        "Zabrze ", "Gliwice ", "Rybnik ", "Jastrzêbie-Zdrój ", "miko³owski ", "¯ory ", _
        "Tychy ", "bieruñsko-lêdziñski ", "Jaworzno ", "Mys³owice ", "Katowice ", _
        "Sosnowiec ", "D¹browa-Górnicza ")).Group.Name = " "
        
         On Error Resume Next
        
        ActiveSheet.Shapes.Range(Array(" ")).Visible = msoFalse
    
        With ActiveSheet.Shapes.Range(Array("Woj")).Glow
        
            .color.ObjectThemeColor = msoThemeColorBackground1
            ' .Transparency = 0.2
            .Radius = 5
        End With
              
        ActiveSheet.Shapes.Range(Array("pucki", "Gdynia", "Sopot", "Gdañsk", "S³upsk", "wejherowski", "lêborski", "s³upski" _
            , "bytowski", "cz³uchowski", "chojnicki", "koœcierski", "kartuski", "gdañski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzyñski")).Group.Name = "pomorskie"
            
        With ActiveSheet.Shapes.Range(Array("pucki", "Gdynia", "Sopot", "Gdañsk", "S³upsk", "wejherowski", "lêborski", "s³upski" _
            , "bytowski", "cz³uchowski", "chojnicki", "koœcierski", "kartuski", "gdañski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzyñski")).Fill
            
            .ForeColor.RGB = RGB(255, 192, 0)
            .Solid
        End With
        
        With ActiveSheet.Shapes.Range(Array("pucki", "Gdynia", "Sopot", "Gdañsk", "S³upsk", "wejherowski", "lêborski", "s³upski" _
            , "bytowski", "cz³uchowski", "chojnicki", "koœcierski", "kartuski", "gdañski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzyñski")).Line
            
            .ForeColor.RGB = RGB(255, 192, 0)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("s³awieñski", "Szczecin", "Koszalin", "bia³ogardzki", "koszaliñski", _
            "szczecinecki", "wa³ecki", "drawski", "œwidwiñski", "ko³obrzeski", "gryficki", _
            "³obeski", "Œwinoujœcie", "goleniowski", "stargardzki", "choszczeñski", _
            "pyrzycki", "myœliborski", "gryfiñski", "policki", "kamieñski")).Group.Name = "zachodniopomorskie"
    
        With ActiveSheet.Shapes.Range(Array("s³awieñski", "Szczecin", "Koszalin", "bia³ogardzki", "koszaliñski", _
            "szczecinecki", "wa³ecki", "drawski", "œwidwiñski", "ko³obrzeski", "gryficki", _
            "³obeski", "Œwinoujœcie", "goleniowski", "stargardzki", "choszczeñski", _
            "pyrzycki", "myœliborski", "gryfiñski", "policki", "kamieñski")).Fill
            .ForeColor.RGB = RGB(192, 0, 0)
        End With

        With ActiveSheet.Shapes.Range(Array("s³awieñski", "Szczecin", "Koszalin", "bia³ogardzki", "koszaliñski", _
            "szczecinecki", "wa³ecki", "drawski", "œwidwiñski", "ko³obrzeski", "gryficki", _
            "³obeski", "Œwinoujœcie", "goleniowski", "stargardzki", "choszczeñski", _
            "pyrzycki", "myœliborski", "gryfiñski", "policki", "kamieñski")).Line
            .ForeColor.RGB = RGB(192, 0, 0)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Elbl¹g", "Olsztyn", "braniewski", "elbl¹ski", "i³awski", _
            "nowomiejski", "dzia³dowski", "ostródzki", "lidzbarski", "bartoszycki", _
            "olsztyñski", "nidzicki", "szczycieñski", "mr¹gowski", "kêtrzyñski", _
            "wêgorzewski", "go³dapski", "olecki", "gi¿ycki", "e³cki", "piski")).Group.Name = "warmiñsko-mazurskie"
    
        With ActiveSheet.Shapes.Range(Array("Elbl¹g", "Olsztyn", "braniewski", "elbl¹ski", "i³awski", _
            "nowomiejski", "dzia³dowski", "ostródzki", "lidzbarski", "bartoszycki", _
            "olsztyñski", "nidzicki", "szczycieñski", "mr¹gowski", "kêtrzyñski", _
            "wêgorzewski", "go³dapski", "olecki", "gi¿ycki", "e³cki", "piski")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent3
        End With
        
        With ActiveSheet.Shapes.Range(Array("Elbl¹g", "Olsztyn", "braniewski", "elbl¹ski", "i³awski", _
            "nowomiejski", "dzia³dowski", "ostródzki", "lidzbarski", "bartoszycki", _
            "olsztyñski", "nidzicki", "szczycieñski", "mr¹gowski", "kêtrzyñski", _
            "wêgorzewski", "go³dapski", "olecki", "gi¿ycki", "e³cki", "piski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent3
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("sêpoleñski", "tucholski", "Bydgoszcz", "Toruñ", "Grudzi¹dz", "W³oc³awek", "che³miñski", _
            "œwiecki", "grudzi¹dzki", "w¹brzeski", "brodnicki", "rypiñski", _
            "golubsko-dobrzyñski", "lipnowski", "w³oc³awski", "radziejowski", _
            "aleksandrowski", "toruñski", "inowroc³awski", "mogileñski", "bydgoski", _
            "¿niñski")).Group.Name = "kujawsko-pomorskie"
            
        With ActiveSheet.Shapes.Range(Array("sêpoleñski", "tucholski", "Bydgoszcz", "Toruñ", "Grudzi¹dz", "W³oc³awek", "che³miñski", _
            "œwiecki", "grudzi¹dzki", "w¹brzeski", "brodnicki", "rypiñski", _
            "golubsko-dobrzyñski", "lipnowski", "w³oc³awski", "radziejowski", _
            "aleksandrowski", "toruñski", "inowroc³awski", "mogileñski", "bydgoski", _
            "¿niñski")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent4
        End With
        
        With ActiveSheet.Shapes.Range(Array("sêpoleñski", "tucholski", "Bydgoszcz", "Toruñ", "Grudzi¹dz", "W³oc³awek", "che³miñski", _
            "œwiecki", "grudzi¹dzki", "w¹brzeski", "brodnicki", "rypiñski", _
            "golubsko-dobrzyñski", "lipnowski", "w³oc³awski", "radziejowski", _
            "aleksandrowski", "toruñski", "inowroc³awski", "mogileñski", "bydgoski", _
            "¿niñski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent4
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Poznañ", "Konin", "Kalisz", "Leszno", "z³otowski", "pilski", "nakielski", _
            "chodzieski", "w¹growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "miêdzychodzki", "nowotomyski", "wolsztyñski", "grodziski (wielkopolskie)", _
            "poznañski", "gnieŸnieñski", "kolski", "koniñski", "s³upecki", "wrzesiñski", _
            "œremski", "jarociñski", "œredzki (wielkopolskie)", "leszczyñski", "gostyñski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszyñski", "ostrzeszowski", "kêpiñski", "turecki", "koœciañski")).Group.Name = "wielkopolskie"
    
        With ActiveSheet.Shapes.Range(Array("Poznañ", "Konin", "Kalisz", "Leszno", "z³otowski", "pilski", "nakielski", _
            "chodzieski", "w¹growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "miêdzychodzki", "nowotomyski", "wolsztyñski", "grodziski (wielkopolskie)", _
            "poznañski", "gnieŸnieñski", "kolski", "koniñski", "s³upecki", "wrzesiñski", _
            "œremski", "jarociñski", "œredzki (wielkopolskie)", "leszczyñski", "gostyñski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszyñski", "ostrzeszowski", "kêpiñski", "turecki", "koœciañski")).Fill
            .ForeColor.ObjectThemeColor = msoThemeColorAccent5
        End With
        
        With ActiveSheet.Shapes.Range(Array("Poznañ", "Konin", "Kalisz", "Leszno", "z³otowski", "pilski", "nakielski", _
            "chodzieski", "w¹growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "miêdzychodzki", "nowotomyski", "wolsztyñski", "grodziski (wielkopolskie)", _
            "poznañski", "gnieŸnieñski", "kolski", "koniñski", "s³upecki", "wrzesiñski", _
            "œremski", "jarociñski", "œredzki (wielkopolskie)", "leszczyñski", "gostyñski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszyñski", "ostrzeszowski", "kêpiñski", "turecki", "koœciañski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent5
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "Gorzów Wielkopolski", "Zielona Góra", "gorzowski", _
            "sulêciñski", "s³ubicki", "miêdzyrzecki", "œwiebodziñski", "zielonogórski", _
            "kroœnieñski (lubuskie)", "¿arski", "nowosolski", "¿agañski", "wschowski")). _
            Group.Name = "lubuskie"
        
        With ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "Gorzów Wielkopolski", "Zielona Góra", "gorzowski", _
            "sulêciñski", "s³ubicki", "miêdzyrzecki", "œwiebodziñski", "zielonogórski", _
            "kroœnieñski (lubuskie)", "¿arski", "nowosolski", "¿agañski", "wschowski")).Fill
            
            .ForeColor.RGB = RGB(165, 27, 96)
        End With
        
        With ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "Gorzów Wielkopolski", "Zielona Góra", "gorzowski", _
            "sulêciñski", "s³ubicki", "miêdzyrzecki", "œwiebodziñski", "zielonogórski", _
            "kroœnieñski (lubuskie)", "¿arski", "nowosolski", "¿agañski", "wschowski")).Line
            .ForeColor.RGB = RGB(165, 27, 96)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Wroc³aw", "Jelenia Góra", "Legnica", "wa³brzyski", "kamiennogórski", "z³otoryjski" _
            , "lubañski", "zgorzelecki", "boles³awiecki", "polkowicki", "g³ogowski", _
            "górowski", "wo³owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "oleœnicki", "wroc³awski", "o³awski", "strzeliñski", "dzier¿oniowski", _
            "z¹bkowicki", "lubiñski", "œredzki (dolnyœl¹sk)", "œwidnicki (dolnoœl¹skie)", _
            "lwówecki", "jeleniogórski", "k³odzki")).Group.Name = "dolnoœl¹skie"
    
        With ActiveSheet.Shapes.Range(Array("Wroc³aw", "Jelenia Góra", "Legnica", "wa³brzyski", "kamiennogórski", "z³otoryjski" _
            , "lubañski", "zgorzelecki", "boles³awiecki", "polkowicki", "g³ogowski", _
            "górowski", "wo³owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "oleœnicki", "wroc³awski", "o³awski", "strzeliñski", "dzier¿oniowski", _
            "z¹bkowicki", "lubiñski", "œredzki (dolnyœl¹sk)", "œwidnicki (dolnoœl¹skie)", _
            "lwówecki", "jeleniogórski", "k³odzki")).Fill
            
            .ForeColor.RGB = RGB(255, 0, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Wroc³aw", "Jelenia Góra", "Legnica", "wa³brzyski", "kamiennogórski", "z³otoryjski" _
            , "lubañski", "zgorzelecki", "boles³awiecki", "polkowicki", "g³ogowski", _
            "górowski", "wo³owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "oleœnicki", "wroc³awski", "o³awski", "strzeliñski", "dzier¿oniowski", _
            "z¹bkowicki", "lubiñski", "œredzki (dolnyœl¹sk)", "œwidnicki (dolnoœl¹skie)", _
            "lwówecki", "jeleniogórski", "k³odzki")).Line
            
            .ForeColor.RGB = RGB(255, 0, 0)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Suwa³ki", "Bia³ystok", "£om¿a", "suwalski", "sejneñski", "augustowski", _
            "grajewski", "kolneñski", "moniecki", "sokólski", "bia³ostocki", "³om¿yñski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Group.Name = "podlaskie"
        
        With ActiveSheet.Shapes.Range(Array("Suwa³ki", "Bia³ystok", "£om¿a", "suwalski", "sejneñski", "augustowski", _
            "grajewski", "kolneñski", "moniecki", "sokólski", "bia³ostocki", "³om¿yñski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        End With
        
        With ActiveSheet.Shapes.Range(Array("Suwa³ki", "Bia³ystok", "£om¿a", "suwalski", "sejneñski", "augustowski", _
            "grajewski", "kolneñski", "moniecki", "sokólski", "bia³ostocki", "³om¿yñski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Ostro³êka", "P³ock", "Warszawa", "Radom", "Siedlce", "ostro³êcki", "przasnyski", "ciechanowski", _
            "m³awski", "¿uromiñski", "sierpecki", "gostyniñski", "p³ocki", "p³oñski", _
            "pu³tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko³owski", "³osicki", "siedlecki", "wêgrowski", _
            "wo³omiñski", "miñski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "¿yrardowski", "piaseczyñski", "otwocki", "grójecki", "przysuski", _
            "bia³obrzeski", "kozienicki", "garwoliñski", "zwoleñski", "szyd³owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Group.Name = "mazowieckie"
            
        With ActiveSheet.Shapes.Range(Array("Ostro³êka", "P³ock", "Warszawa", "Radom", "Siedlce", "ostro³êcki", "przasnyski", "ciechanowski", _
            "m³awski", "¿uromiñski", "sierpecki", "gostyniñski", "p³ocki", "p³oñski", _
            "pu³tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko³owski", "³osicki", "siedlecki", "wêgrowski", _
            "wo³omiñski", "miñski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "¿yrardowski", "piaseczyñski", "otwocki", "grójecki", "przysuski", _
            "bia³obrzeski", "kozienicki", "garwoliñski", "zwoleñski", "szyd³owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Fill
            
            .ForeColor.RGB = RGB(255, 204, 51)
         End With
         
        With ActiveSheet.Shapes.Range(Array("Ostro³êka", "P³ock", "Warszawa", "Radom", "Siedlce", "ostro³êcki", "przasnyski", "ciechanowski", _
            "m³awski", "¿uromiñski", "sierpecki", "gostyniñski", "p³ocki", "p³oñski", _
            "pu³tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko³owski", "³osicki", "siedlecki", "wêgrowski", _
            "wo³omiñski", "miñski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "¿yrardowski", "piaseczyñski", "otwocki", "grójecki", "przysuski", _
            "bia³obrzeski", "kozienicki", "garwoliñski", "zwoleñski", "szyd³owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Line
            
            .ForeColor.RGB = RGB(255, 204, 51)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("rawski", "£ódŸ", "Piotrków Trybunalski", "Skierniewice", "skierniewicki", "³owicki", _
            "kutnowski", "³êczycki", "zgierski", "brzeziñski", "tomaszowski (³ódzkie)", _
            "opoczyñski", "³ódzki wschodni", "piotrkowski", "radomszczañski", "pajêczañski" _
            , "be³chatowski", "pabianicki", "³aski", "zduñskowolski", "poddêbicki", _
            "sieradzki", "wieluñski", "wieruszowski")).Group.Name = "³ódzkie"
        
        With ActiveSheet.Shapes.Range(Array("rawski", "£ódŸ", "Piotrków Trybunalski", "Skierniewice", "skierniewicki", "³owicki", _
            "kutnowski", "³êczycki", "zgierski", "brzeziñski", "tomaszowski (³ódzkie)", _
            "opoczyñski", "³ódzki wschodni", "piotrkowski", "radomszczañski", "pajêczañski" _
            , "be³chatowski", "pabianicki", "³aski", "zduñskowolski", "poddêbicki", _
            "sieradzki", "wieluñski", "wieruszowski")).Fill
            .ForeColor.RGB = RGB(0, 176, 80)
        End With
        
        With ActiveSheet.Shapes.Range(Array("rawski", "£ódŸ", "Piotrków Trybunalski", "Skierniewice", "skierniewicki", "³owicki", _
            "kutnowski", "³êczycki", "zgierski", "brzeziñski", "tomaszowski (³ódzkie)", _
            "opoczyñski", "³ódzki wschodni", "piotrkowski", "radomszczañski", "pajêczañski" _
            , "be³chatowski", "pabianicki", "³aski", "zduñskowolski", "poddêbicki", _
            "sieradzki", "wieluñski", "wieruszowski")).Line
            
            .ForeColor.RGB = RGB(0, 176, 80)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Opole", "g³ubczycki", "kêdzierzyñsko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys³owski", "brzeski (opolskie)", "nyski")).Group.Name = "opolskie"
        
        With ActiveSheet.Shapes.Range(Array("Opole", "g³ubczycki", "kêdzierzyñsko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys³owski", "brzeski (opolskie)", "nyski")).Fill
            
            .ForeColor.RGB = RGB(0, 0, 204)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Opole", "g³ubczycki", "kêdzierzyñsko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys³owski", "brzeski (opolskie)", "nyski")).Line
            
            .ForeColor.RGB = RGB(0, 0, 204)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Czêstochowa", "Bielsko-Bia³a", "Ruda Œl¹ska", "Œwiêtoch³owice", "Chorzów" _
        , "Siemianowice Œl¹skie", "Piekary Œl¹skie", "Bytom", "Zabrze", "Gliwice", "Rybnik", "Jastrzêbie-Zdrój", "¯ory", _
        "Tychy", "Jaworzno", "Mys³owice", "Katowice", "Sosnowiec", "D¹browa-Górnicza", "bieruñsko-lêdziñski", "pszczyñski", _
        "bielski (œl¹skie)", "cieszyñski", "wodzis³awski", "rybnicki", "raciborski", _
        "gliwicki", "miko³owski", "bêdziñski", "tarnogórski", "zawierciañski", _
        "myszkowski", "czêstochowski", "lubliniecki", "k³obucki", "¿ywiecki")).Group.Name = "œl¹skie"
    
        With ActiveSheet.Shapes.Range(Array("Czêstochowa", "Bielsko-Bia³a", "Ruda Œl¹ska", "Œwiêtoch³owice", "Chorzów" _
        , "Siemianowice Œl¹skie", "Piekary Œl¹skie", "Bytom", "Zabrze", "Gliwice", "Rybnik", "Jastrzêbie-Zdrój", "¯ory", _
        "Tychy", "Jaworzno", "Mys³owice", "Katowice", "Sosnowiec", "D¹browa-Górnicza", "bieruñsko-lêdziñski", "pszczyñski", _
        "bielski (œl¹skie)", "cieszyñski", "wodzis³awski", "rybnicki", "raciborski", _
        "gliwicki", "miko³owski", "bêdziñski", "tarnogórski", "zawierciañski", _
        "myszkowski", "czêstochowski", "lubliniecki", "k³obucki", "¿ywiecki")).Fill
            .ForeColor.RGB = RGB(204, 153, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Czêstochowa", "Bielsko-Bia³a", "Ruda Œl¹ska", "Œwiêtoch³owice", "Chorzów" _
        , "Siemianowice Œl¹skie", "Piekary Œl¹skie", "Bytom", "Zabrze", "Gliwice", "Rybnik", "Jastrzêbie-Zdrój", "¯ory", _
        "Tychy", "Jaworzno", "Mys³owice", "Katowice", "Sosnowiec", "D¹browa-Górnicza", "bieruñsko-lêdziñski", "pszczyñski", _
        "bielski (œl¹skie)", "cieszyñski", "wodzis³awski", "rybnicki", "raciborski", _
        "gliwicki", "miko³owski", "bêdziñski", "tarnogórski", "zawierciañski", _
        "myszkowski", "czêstochowski", "lubliniecki", "k³obucki", "¿ywiecki")).Line
             .ForeColor.RGB = RGB(204, 153, 0)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Kielce", "konecki", "skar¿yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w³oszczowski", "jêdrzejowski", "piñczowski", "buski", "kazimierski")).Group.Name = "œwiêtokrzyskie"
            
        With ActiveSheet.Shapes.Range(Array("Kielce", "konecki", "skar¿yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w³oszczowski", "jêdrzejowski", "piñczowski", "buski", "kazimierski")).Fill
            
            .ForeColor.RGB = RGB(204, 51, 255)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Kielce", "konecki", "skar¿yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w³oszczowski", "jêdrzejowski", "piñczowski", "buski", "kazimierski")).Line
            
            .ForeColor.RGB = RGB(204, 51, 255)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Zamoœæ", "Che³m", "Bia³a Podlaska", "Lublin", "bialski", "radzyñski", "³ukowski", "rycki", _
            "parczewski", "w³odawski", "lubartowski", "pu³awski", "opolski (lubelskie)", _
            "lubelski", "kraœnicki", "œwidnicki (lubelskie)", "che³mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi³gorajski", "³êczyñski", _
            "tomaszowski (lubelskie)")).Group.Name = "lubelskie"
            
        With ActiveSheet.Shapes.Range(Array("Zamoœæ", "Che³m", "Bia³a Podlaska", "Lublin", "bialski", "radzyñski", "³ukowski", "rycki", _
            "parczewski", "w³odawski", "lubartowski", "pu³awski", "opolski (lubelskie)", _
            "lubelski", "kraœnicki", "œwidnicki (lubelskie)", "che³mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi³gorajski", "³êczyñski", _
            "tomaszowski (lubelskie)")).Fill
            .ForeColor.RGB = RGB(102, 0, 204)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Zamoœæ", "Che³m", "Bia³a Podlaska", "Lublin", "bialski", "radzyñski", "³ukowski", "rycki", _
            "parczewski", "w³odawski", "lubartowski", "pu³awski", "opolski (lubelskie)", _
            "lubelski", "kraœnicki", "œwidnicki (lubelskie)", "che³mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi³gorajski", "³êczyñski", _
            "tomaszowski (lubelskie)")).Line
            .ForeColor.RGB = RGB(102, 0, 204)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Rzeszów", "Przemyœl", "Krosno", "Tarnobrzeg", "jaros³awski", "przemyski", "le¿ajski", _
            "przeworski", "³añcucki", "ni¿añski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-sêdziszowski", "mielecki", "dêbicki", _
            "brzozowski", "kroœnieñski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy¿owski", "lubaczowski")).Group.Name = "podkarpackie"
        
        With ActiveSheet.Shapes.Range(Array("Rzeszów", "Przemyœl", "Krosno", "Tarnobrzeg", "jaros³awski", "przemyski", "le¿ajski", _
            "przeworski", "³añcucki", "ni¿añski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-sêdziszowski", "mielecki", "dêbicki", _
            "brzozowski", "kroœnieñski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy¿owski", "lubaczowski")).Fill
            
            .ForeColor.RGB = RGB(255, 64, 64)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Rzeszów", "Przemyœl", "Krosno", "Tarnobrzeg", "jaros³awski", "przemyski", "le¿ajski", _
            "przeworski", "³añcucki", "ni¿añski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-sêdziszowski", "mielecki", "dêbicki", _
            "brzozowski", "kroœnieñski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy¿owski", "lubaczowski")).Line
            
            .ForeColor.RGB = RGB(255, 64, 64)
            .Weight = 2.5
        End With

        ActiveSheet.Shapes.Range(Array("Kraków", "Tarnów", "Nowy S¹cz", "d¹browski", "gorlicki", "nowos¹decki", _
            "bocheñski", "brzeski (ma³opolskie)", "tarnowski", "limanowski", "wielicki", _
            "myœlenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "oœwiêcimski", "chrzanowski", "okulski", "nowotarski", "tatrzañski")).Group.Name = "ma³opolskie"
                
        With ActiveSheet.Shapes.Range(Array("Kraków", "Tarnów", "Nowy S¹cz", "d¹browski", "gorlicki", "nowos¹decki", _
            "bocheñski", "brzeski (ma³opolskie)", "tarnowski", "limanowski", "wielicki", _
            "myœlenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "oœwiêcimski", "chrzanowski", "okulski", "nowotarski", "tatrzañski")).Fill
            
            .ForeColor.RGB = RGB(204, 51, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Kraków", "Tarnów", "Nowy S¹cz", "d¹browski", "gorlicki", "nowos¹decki", _
            "bocheñski", "brzeski (ma³opolskie)", "tarnowski", "limanowski", "wielicki", _
            "myœlenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "oœwiêcimski", "chrzanowski", "okulski", "nowotarski", "tatrzañski")).Line
            
            .ForeColor.RGB = RGB(204, 51, 0)
            .Weight = 2.5
        End With
         
        'Wykres 2011
        With ActiveSheet.ChartObjects("Wykres 399")

            With .Chart.SeriesCollection(1)
               
               .Values = Worksheets("Mapka dane").Range("C" & 426 & ":E" & 426)
            
            End With
    
        End With
    
        'Wykres 2035
        With ActiveSheet.ChartObjects("Wykres 784")

            With .Chart.SeriesCollection(1)

                .Values = Worksheets("Mapka dane").Range("BW" & 426 & ":BY" & 426)
        
            End With
        
        End With
    
        'Wykres dynamiki
        With ActiveSheet.ChartObjects("Wykres 1").Chart
    
            .ChartTitle.Text = "Województwo mazowieckie - prognoza liczby ludnoœci na lata 2011-2035"
            .SeriesCollection(1).Values = Worksheets("Mapka dane").Range("C" & 426 & "," & "F" & 426 & "," & "I" & 426 & "," & "L" & 426 & "," & "O" & 426 & "," & "R" & 426 & "," & "U" & 426 & "," & "X" & 426 & "," & "AA" & 426 & "," & "AD" & 426 & "," & "AG" & 426 & "," & "AJ" & 426 & "," & "AM" & 426 & "," & "AP" & 426 & "," & "AS" & 426 & "," & "AV" & 426 & "," & "AY" & 426 & "," & "BB" & 426 & "," & "BE" & 426 & "," & "BK" & 426 & "," & "BN" & 426 & "," & "BQ" & 426 & "," & "BT" & 426 & "," & "BW" & 426)
            .SeriesCollection(2).Values = Worksheets("Mapka dane").Range("D" & 426 & "," & "G" & 426 & "," & "J" & 426 & "," & "M" & 426 & "," & "P" & 426 & "," & "S" & 426 & "," & "V" & 426 & "," & "Y" & 426 & "," & "AB" & 426 & "," & "AE" & 426 & "," & "AH" & 426 & "," & "AK" & 426 & "," & "AN" & 426 & "," & "AQ" & 426 & "," & "AT" & 426 & "," & "AW" & 426 & "," & "AZ" & 426 & "," & "BC" & 426 & "," & "BF" & 426 & "," & "BL" & 426 & "," & "BO" & 426 & "," & "BR" & 426 & "," & "BU" & 426 & "," & "BX" & 426)
            .SeriesCollection(3).Values = Worksheets("Mapka dane").Range("E" & 426 & "," & "H" & 426 & "," & "K" & 426 & "," & "N" & 426 & "," & "Q" & 426 & "," & "T" & 426 & "," & "W" & 426 & "," & "Z" & 426 & "," & "AC" & 426 & "," & "AF" & 426 & "," & "AI" & 426 & "," & "AL" & 426 & "," & "AO" & 426 & "," & "AR" & 426 & "," & "AU" & 426 & "," & "AX" & 426 & "," & "BA" & 426 & "," & "BD" & 426 & "," & "BG" & 426 & "," & "BM" & 426 & "," & "BP" & 426 & "," & "BS" & 426 & "," & "BV" & 426 & "," & "BY" & 426)
    
        End With
        
    End If
    
    RefreshCharts  'odswieza wykresy, ¿eby update'owa³y siê w czasie dzia³ania excela

End Sub


Sub RefreshCharts() ' nie wiem dlaczego, ale to musi byc sub i musi byc num = 57

num = 57

        'Wykres 2011
        With ActiveSheet.ChartObjects("Wykres 399")

            With .Chart.SeriesCollection(1)

               .Values = Worksheets("Powiaty").Range("C" & num & ":E" & num)

            End With

        End With

        'Wykres 2035
        With ActiveSheet.ChartObjects("Wykres 784")

            With .Chart.SeriesCollection(1)

                .Values = Worksheets("Powiaty").Range("BW" & num & ":BY" & num)

            End With

        End With

        'Wykres dynamiki
        With ActiveSheet.ChartObjects("Wykres 1").Chart

            '.ChartTitle.Text = title
            .SeriesCollection(1).Values = Worksheets("Powiaty").Range("C" & num & "," & "F" & num & "," & "I" & num & "," & "L" & num & "," & "O" & num & "," & "R" & num & "," & "U" & num & "," & "X" & num & "," & "AA" & num & "," & "AD" & num & "," & "AG" & num & "," & "AJ" & num & "," & "AM" & num & "," & "AP" & num & "," & "AS" & num & "," & "AV" & num & "," & "AY" & num & "," & "BB" & num & "," & "BE" & num & "," & "BK" & num & "," & "BN" & num & "," & "BQ" & num & "," & "BT" & num & "," & "BW" & num)
            .SeriesCollection(2).Values = Worksheets("Powiaty").Range("D" & num & "," & "G" & num & "," & "J" & num & "," & "M" & num & "," & "P" & num & "," & "S" & num & "," & "V" & num & "," & "Y" & num & "," & "AB" & num & "," & "AE" & num & "," & "AH" & num & "," & "AK" & num & "," & "AN" & num & "," & "AQ" & num & "," & "AT" & num & "," & "AW" & num & "," & "AZ" & num & "," & "BC" & num & "," & "BF" & num & "," & "BL" & num & "," & "BO" & num & "," & "BR" & num & "," & "BU" & num & "," & "BX" & num)
            .SeriesCollection(3).Values = Worksheets("Powiaty").Range("E" & num & "," & "H" & num & "," & "K" & num & "," & "N" & num & "," & "Q" & num & "," & "T" & num & "," & "W" & num & "," & "Z" & num & "," & "AC" & num & "," & "AF" & num & "," & "AI" & num & "," & "AL" & num & "," & "AO" & num & "," & "AR" & num & "," & "AU" & num & "," & "AX" & num & "," & "BA" & num & "," & "BD" & num & "," & "BG" & num & "," & "BM" & num & "," & "BP" & num & "," & "BS" & num & "," & "BV" & num & "," & "BY" & num)

        End With

End Sub
