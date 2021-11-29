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
        
        ActiveSheet.Shapes.Range(Array("Ruda �l�ska ", "�wi�toch�owice ", "Chorz�w " _
        , "b�dzi�ski ", "Siemianowice �l�skie ", "Piekary �l�skie ", "Bytom ", _
        "Zabrze ", "Gliwice ", "Rybnik ", "Jastrz�bie-Zdr�j ", "miko�owski ", "�ory ", _
        "Tychy ", "bieru�sko-l�dzi�ski ", "Jaworzno ", "Mys�owice ", "Katowice ", _
        "Sosnowiec ", "D�browa-G�rnicza ")).Group.Name = " "
        
        On Error Resume Next
    
        ActiveSheet.Shapes.Range(Array(" ")).Visible = msoFalse
          
         With ActiveSheet.Shapes.Range(Array( _
        "hajnowski", "Toru�", "mi�dzyrzecki", "sul�ci�ski", "pucki", "wejherowski", "l�borski", "s�upski", "s�awie�ski", "Gdynia", "S�upsk", "kartuski", "suwalski", "bytowski", "Gda�sk", "braniewski", "go�dapski", "bartoszycki", "w�gorzewski", "sejne�ski", "k�trzy�ski", "koszali�ski", "nowodworski (pomorskie)", _
        "gda�ski", "elbl�ski", "Elbl�g", "olecki", "ko�obrzeski", "Koszalin", "lidzbarski", "Suwa�ki", "gryficki", "gi�ycki", "ko�cierski", "malborski", "bia�ogardzki", "tczewski", "starogardzki", "kamie�ski", "olszty�ski", "augustowski", "�winouj�cie", "ostr�dzki", "e�cki", "chojnicki", "sztumski", "szczecinecki", "cz�uchowski", "�widwi�ski", "mr�gowski", _
        "piski", "�obeski", "i�awski", "kwidzy�ski", "goleniowski", "Olsztyn", "grajewski", "policki", "szczycie�ski", "sok�lski", "tucholski", "moniecki", "drawski", "�wiecki", "z�otowski", "s�pole�ski", "grudzi�dzki", "Szczecin", "stargardzki", "kolne�ski", "nowomiejski", "Grudzi�dz", "nidzicki", "ostro��cki", "�om�y�ski", "wa�ecki", "bia�ostocki", "brodnicki", "bydgoski", "dzia�dowski", "gryfi�ski", "che�mi�ski", "w�brzeski", "choszcze�ski", _
        "przasnyski", "pilski", "pyrzycki", "nakielski", "m�awski", "toru�ski", "Bia�ystok", "golubsko-dobrzy�ski", "zambrowski", "�om�a", "rypi�ski", "Bydgoszcz", "�uromi�ski", "wysokomazowiecki", "strzelecko-drezdenecki", "czarnkowsko-trzcianecki", "makowski", "Ostro��ka", "ciechanowski", "my�liborski", "chodzieski")) _
        .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    
        With ActiveSheet.Shapes.Range(Array( _
        "hajnowski", "Toru�", "mi�dzyrzecki", "sul�ci�ski", "pucki", "wejherowski", "l�borski", "s�upski", "s�awie�ski", "Gdynia", "S�upsk", "kartuski", "suwalski", "bytowski", "Gda�sk", "braniewski", "go�dapski", "bartoszycki", "w�gorzewski", "sejne�ski", "k�trzy�ski", "koszali�ski", "nowodworski (pomorskie)", _
        "gda�ski", "elbl�ski", "Elbl�g", "olecki", "ko�obrzeski", "Koszalin", "lidzbarski", "Suwa�ki", "gryficki", "gi�ycki", "ko�cierski", "malborski", "bia�ogardzki", "tczewski", "starogardzki", "kamie�ski", "olszty�ski", "augustowski", "�winouj�cie", "ostr�dzki", "e�cki", "chojnicki", "sztumski", "szczecinecki", "cz�uchowski", "�widwi�ski", "mr�gowski", _
        "piski", "�obeski", "i�awski", "kwidzy�ski", "goleniowski", "Olsztyn", "grajewski", "policki", "szczycie�ski", "sok�lski", "tucholski", "moniecki", "drawski", "�wiecki", "z�otowski", "s�pole�ski", "grudzi�dzki", "Szczecin", "stargardzki", "kolne�ski", "nowomiejski", "Grudzi�dz", "nidzicki", "ostro��cki", "�om�y�ski", "wa�ecki", "bia�ostocki", "brodnicki", "bydgoski", "dzia�dowski", "gryfi�ski", "che�mi�ski", "w�brzeski", "choszcze�ski", _
        "przasnyski", "pilski", "pyrzycki", "nakielski", "m�awski", "toru�ski", "Bia�ystok", "golubsko-dobrzy�ski", "zambrowski", "�om�a", "rypi�ski", "Bydgoszcz", "�uromi�ski", "wysokomazowiecki", "strzelecko-drezdenecki", "czarnkowsko-trzcianecki", "makowski", "Ostro��ka", "ciechanowski", "my�liborski", "chodzieski")) _
        .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
            .Visible = msoTrue
            .Weight = 2.5
        End With
        
        ActiveSheet.Shapes.Range(Array( _
        "hajnowski", "Toru�", "mi�dzyrzecki", "sul�ci�ski", "pucki", "wejherowski", "l�borski", "s�upski", "s�awie�ski", "Gdynia", "S�upsk", "kartuski", "suwalski", "bytowski", "Gda�sk", "braniewski", "go�dapski", "bartoszycki", "w�gorzewski", "sejne�ski", "k�trzy�ski", "koszali�ski", "nowodworski (pomorskie)", _
        "gda�ski", "elbl�ski", "Elbl�g", "olecki", "ko�obrzeski", "Koszalin", "lidzbarski", "Suwa�ki", "gryficki", "gi�ycki", "ko�cierski", "malborski", "bia�ogardzki", "tczewski", "starogardzki", "kamie�ski", "olszty�ski", "augustowski", "�winouj�cie", "ostr�dzki", "e�cki", "chojnicki", "sztumski", "szczecinecki", "cz�uchowski", "�widwi�ski", "mr�gowski", _
        "piski", "�obeski", "i�awski", "kwidzy�ski", "goleniowski", "Olsztyn", "grajewski", "policki", "szczycie�ski", "sok�lski", "tucholski", "moniecki", "drawski", "�wiecki", "z�otowski", "s�pole�ski", "grudzi�dzki", "Szczecin", "stargardzki", "kolne�ski", "nowomiejski", "Grudzi�dz", "nidzicki", "ostro��cki", "�om�y�ski", "wa�ecki", "bia�ostocki", "brodnicki", "bydgoski", "dzia�dowski", "gryfi�ski", "che�mi�ski", "w�brzeski", "choszcze�ski", _
        "przasnyski", "pilski", "pyrzycki", "nakielski", "m�awski", "toru�ski", "Bia�ystok", "golubsko-dobrzy�ski", "zambrowski", "�om�a", "rypi�ski", "Bydgoszcz", "�uromi�ski", "wysokomazowiecki", "strzelecko-drezdenecki", "czarnkowsko-trzcianecki", "makowski", "Ostro��ka", "ciechanowski", "my�liborski", "chodzieski")) _
        .Group.Name = "Grupa 1"
        
        With ActiveSheet.Shapes.Range(Array("�ni�ski", "w�growiecki", "Konin", "Sopot", "Siedlce", "ostrowski (mazowieckie)", "lipnowski", "bielski (podlaskie)", "inowroc�awski", "sierpecki", "gorzowski", "aleksandrowski", "p�o�ski", "obornicki", "wyszkowski", "mi�dzychodzki", "Gorz�w Wielkopolski", "mogile�ski", "pu�tuski", "szamotulski", "p�ocki", "w�oc�awski", "siemiatycki", "gnie�nie�ski", "W�oc�awek", "soko�owski", "w�growski", "radziejowski", "pozna�ski", "nowodworski (mazowieckie)", "s�ubicki", "P�ock", "wo�omi�ski", "legionowski", "s�upecki", "nowotomyski", "koni�ski", "Pozna�", "gostyni�ski", "siedlecki", "�osicki", "�wiebodzi�ski", "wrzesi�ski", "kolski", "mi�ski", "warszawski zachodni", "kutnowski", "sochaczewski", "Warszawa")) _
         .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    
         With ActiveSheet.Shapes.Range(Array("�ni�ski", "w�growiecki", "Konin", "Sopot", "Siedlce", "ostrowski (mazowieckie)", "lipnowski", "bielski (podlaskie)", "inowroc�awski", "sierpecki", "gorzowski", "aleksandrowski", "p�o�ski", "obornicki", "wyszkowski", "mi�dzychodzki", "Gorz�w Wielkopolski", "mogile�ski", "pu�tuski", "szamotulski", "p�ocki", "w�oc�awski", "siemiatycki", "gnie�nie�ski", "W�oc�awek", "soko�owski", "w�growski", "radziejowski", "pozna�ski", "nowodworski (mazowieckie)", "s�ubicki", "P�ock", "wo�omi�ski", "legionowski", "s�upecki", "nowotomyski", "koni�ski", "Pozna�", "gostyni�ski", "siedlecki", "�osicki", "�wiebodzi�ski", "wrzesi�ski", "kolski", "mi�ski", "warszawski zachodni", "kutnowski", "sochaczewski", "Warszawa")) _
         .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
            .Visible = msoTrue
            .Weight = 2.5
        End With
        
        ActiveSheet.Shapes.Range(Array("�ni�ski", "w�growiecki", "Konin", "Sopot", "Siedlce", "ostrowski (mazowieckie)", "lipnowski", "bielski (podlaskie)", "inowroc�awski", "sierpecki", "gorzowski", "aleksandrowski", "p�o�ski", "obornicki", "wyszkowski", "mi�dzychodzki", "Gorz�w Wielkopolski", "mogile�ski", "pu�tuski", "szamotulski", "p�ocki", "w�oc�awski", "siemiatycki", "gnie�nie�ski", "W�oc�awek", "soko�owski", "w�growski", "radziejowski", "pozna�ski", "nowodworski (mazowieckie)", "s�ubicki", "P�ock", "wo�omi�ski", "legionowski", "s�upecki", "nowotomyski", "koni�ski", "Pozna�", "gostyni�ski", "siedlecki", "�osicki", "�wiebodzi�ski", "wrzesi�ski", "kolski", "mi�ski", "warszawski zachodni", "kutnowski", "sochaczewski", "Warszawa")) _
         .Group.Name = "Grupa 2"
        
        With ActiveSheet.Shapes.Range(Array( _
        "bialski", "�redzki (wielkopolskie)", "grodziski (wielkopolskie)", "�owicki", "kro�nie�ski (lubuskie)", "wolszty�ski", "zielonog�rski", "otwocki", "ko�cia�ski", "pruszkowski", "��czycki", "grodziski (mazowieckie)", "�remski", "turecki", "�yrardowski", "piaseczy�ski", "skierniewicki", "jaroci�ski", "Bia�a Podlaska", "�ukowski", "pleszewski", "podd�bicki", "nowosolski", "garwoli�ski", "kaliski", "Zielona G�ra", "wschowski", "zgierski", "radzy�ski", "leszczy�ski", "Skierniewice", "gosty�ski", "gr�jecki", "brzezi�ski", "�arski", "krotoszy�ski", "rawski", "Leszno", "kozienicki", "w�odawski", "��d�", "parczewski", "sieradzki", "rycki", "g�ogowski", "�aga�ski", "ostrowski (wielkopolskie)", "pabianicki", "g�rowski", "Kalisz", "tomaszowski (��dzkie)", "zdu�skowolski", "�aski", "bia�obrzeski", "lubartowski", "rawicki", "polkowicki", _
        "milicki", "piotrkowski", "przysuski", "lubi�ski", "radomski", "pu�awski", "opoczy�ski", "trzebnicki", "boles�awiecki", "ostrzeszowski", "be�chatowski", "wo�owski", "zgorzelecki", "Radom", "zwole�ski", "lubelski", "��czy�ski", "Piotrk�w Trybunalski", "che�mski", "ole�nicki", "wieruszowski", "legnicki", "wielu�ski", "k�pi�ski", "szyd�owiecki", "Lublin", "konecki", "opolski (lubelskie)", "�widnicki (lubelskie)", "lipski", "z�otoryjski", "paj�cza�ski", "radomszcza�ski", "�redzki (dolny�l�sk)", "wroc�awski", "Legnica", "Che�m", "lw�wecki", "luba�ski", "starachowicki", "Wroc�aw", "skar�yski", "namys�owski", "krasnostawski", "kluczborski", "jaworski", "oleski", "opatowski", "o�awski", "ostrowiecki", "kra�nicki", "k�obucki", "hrubieszowski", "kielecki", "jeleniog�rski", "�widnicki (dolno�l�skie)", "cz�stochowski", "w�oszczowski", "brzeski (opolskie)", "zamojski", _
        "Jelenia G�ra", "opolski (opolskie)", "janowski", "bi�gorajski", "Kielce", "kamiennog�rski", "strzeli�ski", "sandomierski", "wa�brzyski", "Cz�stochowa", "j�drzejowski", "lubliniecki", "stalowowolski", "dzier�oniowski", "Zamo��", "tomaszowski (lubelskie)", "tarnobrzeski", "z�bkowicki", "zawiercia�ski", "staszowski", "Opole", "strzelecki", "myszkowski", "k�odzki", "pi�czowski", "Tarnobrzeg", "ni�a�ski", "nyski", "buski", "tarnog�rski", "krapkowicki", "prudnicki", "gliwicki", "b�dzi�ski", "mielecki", "miechowski", "kolbuszowski", "lubaczowski", "D�browa-G�rnicza", "okulski", "le�ajski", "Bytom", "Piekary �l�skie", "k�dzierzy�sko-kozielski", "rzeszowski", "kazimierski", "Gliwice", "Zabrze", "d�browski", "przeworski", "Siemianowice �l�skie", "Chorz�w", "Ruda �l�ska", "proszowicki", "�wi�toch�owice", "Sosnowiec", "g�ubczycki", "krakowski", "Katowice", "�a�cucki", _
        "jaros�awski", "Mys�owice", "raciborski", "miko�owski", "tarnowski", "chrzanowski", "brzeski (ma�opolskie)", "ropczycko-s�dziszowski", "rybnicki", "d�bicki", "Rybnik", "Tychy", "bieru�sko-l�dzi�ski", "boche�ski", "Krak�w", "Rzesz�w", "o�wi�cimski", "wielicki", "pszczy�ski", "Tarn�w", "wodzis�awski", "wadowicki", "przemyski", "strzy�owski", "bielski (�l�skie)", "cieszy�ski", "my�lenicki", "jasielski", "Bielsko-Bia�a", "kro�nie�ski", "brzozowski", "Przemy�l", "limanowski", "�ywiecki", "suski", "nowos�decki", "gorlicki", "Krosno", "bieszczadzki", "sanocki", "Nowy S�cz", "nowotarski", "leski", "tatrza�ski", "�ory", "Jastrz�bie-Zdr�j", "Jaworzno", "��dzki wschodni")) _
        .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(220, 20, 60)
            .Solid
        End With
    
       With ActiveSheet.Shapes.Range(Array( _
        "bialski", "�redzki (wielkopolskie)", "grodziski (wielkopolskie)", "�owicki", "kro�nie�ski (lubuskie)", "wolszty�ski", "zielonog�rski", "otwocki", "ko�cia�ski", "pruszkowski", "��czycki", "grodziski (mazowieckie)", "�remski", "turecki", "�yrardowski", "piaseczy�ski", "skierniewicki", "jaroci�ski", "Bia�a Podlaska", "�ukowski", "pleszewski", "podd�bicki", "nowosolski", "garwoli�ski", "kaliski", "Zielona G�ra", "wschowski", "zgierski", "radzy�ski", "leszczy�ski", "Skierniewice", "gosty�ski", "gr�jecki", "brzezi�ski", "�arski", "krotoszy�ski", "rawski", "Leszno", "kozienicki", "w�odawski", "��d�", "parczewski", "sieradzki", "rycki", "g�ogowski", "�aga�ski", "ostrowski (wielkopolskie)", "pabianicki", "g�rowski", "Kalisz", "tomaszowski (��dzkie)", "zdu�skowolski", "�aski", "bia�obrzeski", "lubartowski", "rawicki", "polkowicki", _
        "milicki", "piotrkowski", "przysuski", "lubi�ski", "radomski", "pu�awski", "opoczy�ski", "trzebnicki", "boles�awiecki", "ostrzeszowski", "be�chatowski", "wo�owski", "zgorzelecki", "Radom", "zwole�ski", "lubelski", "��czy�ski", "Piotrk�w Trybunalski", "che�mski", "ole�nicki", "wieruszowski", "legnicki", "wielu�ski", "k�pi�ski", "szyd�owiecki", "Lublin", "konecki", "opolski (lubelskie)", "�widnicki (lubelskie)", "lipski", "z�otoryjski", "paj�cza�ski", "radomszcza�ski", "�redzki (dolny�l�sk)", "wroc�awski", "Legnica", "Che�m", "lw�wecki", "luba�ski", "starachowicki", "Wroc�aw", "skar�yski", "namys�owski", "krasnostawski", "kluczborski", "jaworski", "oleski", "opatowski", "o�awski", "ostrowiecki", "kra�nicki", "k�obucki", "hrubieszowski", "kielecki", "jeleniog�rski", "�widnicki (dolno�l�skie)", "cz�stochowski", "w�oszczowski", "brzeski (opolskie)", "zamojski", _
        "Jelenia G�ra", "opolski (opolskie)", "janowski", "bi�gorajski", "Kielce", "kamiennog�rski", "strzeli�ski", "sandomierski", "wa�brzyski", "Cz�stochowa", "j�drzejowski", "lubliniecki", "stalowowolski", "dzier�oniowski", "Zamo��", "tomaszowski (lubelskie)", "tarnobrzeski", "z�bkowicki", "zawiercia�ski", "staszowski", "Opole", "strzelecki", "myszkowski", "k�odzki", "pi�czowski", "Tarnobrzeg", "ni�a�ski", "nyski", "buski", "tarnog�rski", "krapkowicki", "prudnicki", "gliwicki", "b�dzi�ski", "mielecki", "miechowski", "kolbuszowski", "lubaczowski", "D�browa-G�rnicza", "okulski", "le�ajski", "Bytom", "Piekary �l�skie", "k�dzierzy�sko-kozielski", "rzeszowski", "kazimierski", "Gliwice", "Zabrze", "d�browski", "przeworski", "Siemianowice �l�skie", "Chorz�w", "Ruda �l�ska", "proszowicki", "�wi�toch�owice", "Sosnowiec", "g�ubczycki", "krakowski", "Katowice", "�a�cucki", _
        "jaros�awski", "Mys�owice", "raciborski", "miko�owski", "tarnowski", "chrzanowski", "brzeski (ma�opolskie)", "ropczycko-s�dziszowski", "rybnicki", "d�bicki", "Rybnik", "Tychy", "bieru�sko-l�dzi�ski", "boche�ski", "Krak�w", "Rzesz�w", "o�wi�cimski", "wielicki", "pszczy�ski", "Tarn�w", "wodzis�awski", "wadowicki", "przemyski", "strzy�owski", "bielski (�l�skie)", "cieszy�ski", "my�lenicki", "jasielski", "Bielsko-Bia�a", "kro�nie�ski", "brzozowski", "Przemy�l", "limanowski", "�ywiecki", "suski", "nowos�decki", "gorlicki", "Krosno", "bieszczadzki", "sanocki", "Nowy S�cz", "nowotarski", "leski", "tatrza�ski", "�ory", "Jastrz�bie-Zdr�j", "Jaworzno", "��dzki wschodni")) _
        .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(220, 20, 60)
            .Visible = msoTrue
            .Weight = 2.5
        End With
        
        ActiveSheet.Shapes.Range(Array( _
        "bialski", "�redzki (wielkopolskie)", "grodziski (wielkopolskie)", "�owicki", "kro�nie�ski (lubuskie)", "wolszty�ski", "zielonog�rski", "otwocki", "ko�cia�ski", "pruszkowski", "��czycki", "grodziski (mazowieckie)", "�remski", "turecki", "�yrardowski", "piaseczy�ski", "skierniewicki", "jaroci�ski", "Bia�a Podlaska", "�ukowski", "pleszewski", "podd�bicki", "nowosolski", "garwoli�ski", "kaliski", "Zielona G�ra", "wschowski", "zgierski", "radzy�ski", "leszczy�ski", "Skierniewice", "gosty�ski", "gr�jecki", "brzezi�ski", "�arski", "krotoszy�ski", "rawski", "Leszno", "kozienicki", "w�odawski", "��d�", "parczewski", "sieradzki", "rycki", "g�ogowski", "�aga�ski", "ostrowski (wielkopolskie)", "pabianicki", "g�rowski", "Kalisz", "tomaszowski (��dzkie)", "zdu�skowolski", "�aski", "bia�obrzeski", "lubartowski", "rawicki", "polkowicki", _
        "milicki", "piotrkowski", "przysuski", "lubi�ski", "radomski", "pu�awski", "opoczy�ski", "trzebnicki", "boles�awiecki", "ostrzeszowski", "be�chatowski", "wo�owski", "zgorzelecki", "Radom", "zwole�ski", "lubelski", "��czy�ski", "Piotrk�w Trybunalski", "che�mski", "ole�nicki", "wieruszowski", "legnicki", "wielu�ski", "k�pi�ski", "szyd�owiecki", "Lublin", "konecki", "opolski (lubelskie)", "�widnicki (lubelskie)", "lipski", "z�otoryjski", "paj�cza�ski", "radomszcza�ski", "�redzki (dolny�l�sk)", "wroc�awski", "Legnica", "Che�m", "lw�wecki", "luba�ski", "starachowicki", "Wroc�aw", "skar�yski", "namys�owski", "krasnostawski", "kluczborski", "jaworski", "oleski", "opatowski", "o�awski", "ostrowiecki", "kra�nicki", "k�obucki", "hrubieszowski", "kielecki", "jeleniog�rski", "�widnicki (dolno�l�skie)", "cz�stochowski", "w�oszczowski", "brzeski (opolskie)", "zamojski", _
        "Jelenia G�ra", "opolski (opolskie)", "janowski", "bi�gorajski", "Kielce", "kamiennog�rski", "strzeli�ski", "sandomierski", "wa�brzyski", "Cz�stochowa", "j�drzejowski", "lubliniecki", "stalowowolski", "dzier�oniowski", "Zamo��", "tomaszowski (lubelskie)", "tarnobrzeski", "z�bkowicki", "zawiercia�ski", "staszowski", "Opole", "strzelecki", "myszkowski", "k�odzki", "pi�czowski", "Tarnobrzeg", "ni�a�ski", "nyski", "buski", "tarnog�rski", "krapkowicki", "prudnicki", "gliwicki", "b�dzi�ski", "mielecki", "miechowski", "kolbuszowski", "lubaczowski", "D�browa-G�rnicza", "okulski", "le�ajski", "Bytom", "Piekary �l�skie", "k�dzierzy�sko-kozielski", "rzeszowski", "kazimierski", "Gliwice", "Zabrze", "d�browski", "przeworski", "Siemianowice �l�skie", "Chorz�w", "Ruda �l�ska", "proszowicki", "�wi�toch�owice", "Sosnowiec", "g�ubczycki", "krakowski", "Katowice", "�a�cucki", _
        "jaros�awski", "Mys�owice", "raciborski", "miko�owski", "tarnowski", "chrzanowski", "brzeski (ma�opolskie)", "ropczycko-s�dziszowski", "rybnicki", "d�bicki", "Rybnik", "Tychy", "bieru�sko-l�dzi�ski", "boche�ski", "Krak�w", "Rzesz�w", "o�wi�cimski", "wielicki", "pszczy�ski", "Tarn�w", "wodzis�awski", "wadowicki", "przemyski", "strzy�owski", "bielski (�l�skie)", "cieszy�ski", "my�lenicki", "jasielski", "Bielsko-Bia�a", "kro�nie�ski", "brzozowski", "Przemy�l", "limanowski", "�ywiecki", "suski", "nowos�decki", "gorlicki", "Krosno", "bieszczadzki", "sanocki", "Nowy S�cz", "nowotarski", "leski", "tatrza�ski", "�ory", "Jastrz�bie-Zdr�j", "Jaworzno", "��dzki wschodni")) _
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
    
            .ChartTitle.Text = "Polska - prognoza liczby ludno�ci na lata 2011-2035"
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
        ActiveSheet.Shapes.Range(Array("warmi�sko-mazurskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("kujawsko-pomorskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("wielkopolskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("lubuskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("dolno�l�skie")).Ungroup
        ActiveSheet.Shapes.Range(Array("opolskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("�l�skie")).Ungroup
        ActiveSheet.Shapes.Range(Array("mazowieckie")).Ungroup
        ActiveSheet.Shapes.Range(Array("podlaskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("�wi�tokrzyskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("lubelskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("��dzkie")).Ungroup
        ActiveSheet.Shapes.Range(Array("podkarpackie")).Ungroup
        ActiveSheet.Shapes.Range(Array("ma�opolskie")).Ungroup
        
        ActiveSheet.Shapes.Range(Array("Powiat")).Glow.Radius = 0
        ActiveSheet.Shapes.Range(Array("Woj")).Glow.Radius = 0
    
        With ActiveSheet.Shapes.Range(Array("Kraj")).Glow
            .color.ObjectThemeColor = msoThemeColorBackground1
            ' .Transparency = 0.2
            .Radius = 5
        End With
        
        With ActiveSheet.Shapes.Range(Array( _
        "hajnowski", "Toru�", "mi�dzyrzecki", "sul�ci�ski", "pucki", "wejherowski", "l�borski", "s�upski", "s�awie�ski", "Gdynia", "S�upsk", "kartuski", "suwalski", "bytowski", "Gda�sk", "braniewski", "go�dapski", "bartoszycki", "w�gorzewski", "sejne�ski", "k�trzy�ski", "koszali�ski", "nowodworski (pomorskie)", _
        "gda�ski", "elbl�ski", "Elbl�g", "olecki", "ko�obrzeski", "Koszalin", "lidzbarski", "Suwa�ki", "gryficki", "gi�ycki", "ko�cierski", "malborski", "bia�ogardzki", "tczewski", "starogardzki", "kamie�ski", "olszty�ski", "augustowski", "�winouj�cie", "ostr�dzki", "e�cki", "chojnicki", "sztumski", "szczecinecki", "cz�uchowski", "�widwi�ski", "mr�gowski", _
        "piski", "�obeski", "i�awski", "kwidzy�ski", "goleniowski", "Olsztyn", "grajewski", "policki", "szczycie�ski", "sok�lski", "tucholski", "moniecki", "drawski", "�wiecki", "z�otowski", "s�pole�ski", "grudzi�dzki", "Szczecin", "stargardzki", "kolne�ski", "nowomiejski", "Grudzi�dz", "nidzicki", "ostro��cki", "�om�y�ski", "wa�ecki", "bia�ostocki", "brodnicki", "bydgoski", "dzia�dowski", "gryfi�ski", "che�mi�ski", "w�brzeski", "choszcze�ski", _
        "przasnyski", "pilski", "pyrzycki", "nakielski", "m�awski", "toru�ski", "Bia�ystok", "golubsko-dobrzy�ski", "zambrowski", "�om�a", "rypi�ski", "Bydgoszcz", "�uromi�ski", "wysokomazowiecki", "strzelecko-drezdenecki", "czarnkowsko-trzcianecki", "makowski", "Ostro��ka", "ciechanowski", "my�liborski", "chodzieski")) _
        .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    
         With ActiveSheet.Shapes.Range(Array( _
        "hajnowski", "Toru�", "mi�dzyrzecki", "sul�ci�ski", "pucki", "wejherowski", "l�borski", "s�upski", "s�awie�ski", "Gdynia", "S�upsk", "kartuski", "suwalski", "bytowski", "Gda�sk", "braniewski", "go�dapski", "bartoszycki", "w�gorzewski", "sejne�ski", "k�trzy�ski", "koszali�ski", "nowodworski (pomorskie)", _
        "gda�ski", "elbl�ski", "Elbl�g", "olecki", "ko�obrzeski", "Koszalin", "lidzbarski", "Suwa�ki", "gryficki", "gi�ycki", "ko�cierski", "malborski", "bia�ogardzki", "tczewski", "starogardzki", "kamie�ski", "olszty�ski", "augustowski", "�winouj�cie", "ostr�dzki", "e�cki", "chojnicki", "sztumski", "szczecinecki", "cz�uchowski", "�widwi�ski", "mr�gowski", _
        "piski", "�obeski", "i�awski", "kwidzy�ski", "goleniowski", "Olsztyn", "grajewski", "policki", "szczycie�ski", "sok�lski", "tucholski", "moniecki", "drawski", "�wiecki", "z�otowski", "s�pole�ski", "grudzi�dzki", "Szczecin", "stargardzki", "kolne�ski", "nowomiejski", "Grudzi�dz", "nidzicki", "ostro��cki", "�om�y�ski", "wa�ecki", "bia�ostocki", "brodnicki", "bydgoski", "dzia�dowski", "gryfi�ski", "che�mi�ski", "w�brzeski", "choszcze�ski", _
        "przasnyski", "pilski", "pyrzycki", "nakielski", "m�awski", "toru�ski", "Bia�ystok", "golubsko-dobrzy�ski", "zambrowski", "�om�a", "rypi�ski", "Bydgoszcz", "�uromi�ski", "wysokomazowiecki", "strzelecko-drezdenecki", "czarnkowsko-trzcianecki", "makowski", "Ostro��ka", "ciechanowski", "my�liborski", "chodzieski")) _
        .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
            .Visible = msoTrue
            .Weight = 2.5
        End With
        
         ActiveSheet.Shapes.Range(Array( _
        "hajnowski", "Toru�", "mi�dzyrzecki", "sul�ci�ski", "pucki", "wejherowski", "l�borski", "s�upski", "s�awie�ski", "Gdynia", "S�upsk", "kartuski", "suwalski", "bytowski", "Gda�sk", "braniewski", "go�dapski", "bartoszycki", "w�gorzewski", "sejne�ski", "k�trzy�ski", "koszali�ski", "nowodworski (pomorskie)", _
        "gda�ski", "elbl�ski", "Elbl�g", "olecki", "ko�obrzeski", "Koszalin", "lidzbarski", "Suwa�ki", "gryficki", "gi�ycki", "ko�cierski", "malborski", "bia�ogardzki", "tczewski", "starogardzki", "kamie�ski", "olszty�ski", "augustowski", "�winouj�cie", "ostr�dzki", "e�cki", "chojnicki", "sztumski", "szczecinecki", "cz�uchowski", "�widwi�ski", "mr�gowski", _
        "piski", "�obeski", "i�awski", "kwidzy�ski", "goleniowski", "Olsztyn", "grajewski", "policki", "szczycie�ski", "sok�lski", "tucholski", "moniecki", "drawski", "�wiecki", "z�otowski", "s�pole�ski", "grudzi�dzki", "Szczecin", "stargardzki", "kolne�ski", "nowomiejski", "Grudzi�dz", "nidzicki", "ostro��cki", "�om�y�ski", "wa�ecki", "bia�ostocki", "brodnicki", "bydgoski", "dzia�dowski", "gryfi�ski", "che�mi�ski", "w�brzeski", "choszcze�ski", _
        "przasnyski", "pilski", "pyrzycki", "nakielski", "m�awski", "toru�ski", "Bia�ystok", "golubsko-dobrzy�ski", "zambrowski", "�om�a", "rypi�ski", "Bydgoszcz", "�uromi�ski", "wysokomazowiecki", "strzelecko-drezdenecki", "czarnkowsko-trzcianecki", "makowski", "Ostro��ka", "ciechanowski", "my�liborski", "chodzieski")) _
        .Group.Name = "Grupa 1"
        
         With ActiveSheet.Shapes.Range(Array("�ni�ski", "w�growiecki", "Konin", "Sopot", "Siedlce", "ostrowski (mazowieckie)", "lipnowski", "bielski (podlaskie)", "inowroc�awski", "sierpecki", "gorzowski", "aleksandrowski", "p�o�ski", "obornicki", "wyszkowski", "mi�dzychodzki", "Gorz�w Wielkopolski", "mogile�ski", "pu�tuski", "szamotulski", "p�ocki", "w�oc�awski", "siemiatycki", "gnie�nie�ski", "W�oc�awek", "soko�owski", "w�growski", "radziejowski", "pozna�ski", "nowodworski (mazowieckie)", "s�ubicki", "P�ock", "wo�omi�ski", "legionowski", "s�upecki", "nowotomyski", "koni�ski", "Pozna�", "gostyni�ski", "siedlecki", "�osicki", "�wiebodzi�ski", "wrzesi�ski", "kolski", "mi�ski", "warszawski zachodni", "kutnowski", "sochaczewski", "Warszawa")) _
         .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
        End With
    
         With ActiveSheet.Shapes.Range(Array("�ni�ski", "w�growiecki", "Konin", "Sopot", "Siedlce", "ostrowski (mazowieckie)", "lipnowski", "bielski (podlaskie)", "inowroc�awski", "sierpecki", "gorzowski", "aleksandrowski", "p�o�ski", "obornicki", "wyszkowski", "mi�dzychodzki", "Gorz�w Wielkopolski", "mogile�ski", "pu�tuski", "szamotulski", "p�ocki", "w�oc�awski", "siemiatycki", "gnie�nie�ski", "W�oc�awek", "soko�owski", "w�growski", "radziejowski", "pozna�ski", "nowodworski (mazowieckie)", "s�ubicki", "P�ock", "wo�omi�ski", "legionowski", "s�upecki", "nowotomyski", "koni�ski", "Pozna�", "gostyni�ski", "siedlecki", "�osicki", "�wiebodzi�ski", "wrzesi�ski", "kolski", "mi�ski", "warszawski zachodni", "kutnowski", "sochaczewski", "Warszawa")) _
         .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 255, 255)
            .Visible = msoTrue
            .Weight = 2.5
        End With
        
        ActiveSheet.Shapes.Range(Array("�ni�ski", "w�growiecki", "Konin", "Sopot", "Siedlce", "ostrowski (mazowieckie)", "lipnowski", "bielski (podlaskie)", "inowroc�awski", "sierpecki", "gorzowski", "aleksandrowski", "p�o�ski", "obornicki", "wyszkowski", "mi�dzychodzki", "Gorz�w Wielkopolski", "mogile�ski", "pu�tuski", "szamotulski", "p�ocki", "w�oc�awski", "siemiatycki", "gnie�nie�ski", "W�oc�awek", "soko�owski", "w�growski", "radziejowski", "pozna�ski", "nowodworski (mazowieckie)", "s�ubicki", "P�ock", "wo�omi�ski", "legionowski", "s�upecki", "nowotomyski", "koni�ski", "Pozna�", "gostyni�ski", "siedlecki", "�osicki", "�wiebodzi�ski", "wrzesi�ski", "kolski", "mi�ski", "warszawski zachodni", "kutnowski", "sochaczewski", "Warszawa")) _
         .Group.Name = "Grupa 2"
        
       With ActiveSheet.Shapes.Range(Array( _
        "bialski", "�redzki (wielkopolskie)", "grodziski (wielkopolskie)", "�owicki", "kro�nie�ski (lubuskie)", "wolszty�ski", "zielonog�rski", "otwocki", "ko�cia�ski", "pruszkowski", "��czycki", "grodziski (mazowieckie)", "�remski", "turecki", "�yrardowski", "piaseczy�ski", "skierniewicki", "jaroci�ski", "Bia�a Podlaska", "�ukowski", "pleszewski", "podd�bicki", "nowosolski", "garwoli�ski", "kaliski", "Zielona G�ra", "wschowski", "zgierski", "radzy�ski", "leszczy�ski", "Skierniewice", "gosty�ski", "gr�jecki", "brzezi�ski", "�arski", "krotoszy�ski", "rawski", "Leszno", "kozienicki", "w�odawski", "��d�", "parczewski", "sieradzki", "rycki", "g�ogowski", "�aga�ski", "ostrowski (wielkopolskie)", "pabianicki", "g�rowski", "Kalisz", "tomaszowski (��dzkie)", "zdu�skowolski", "�aski", "bia�obrzeski", "lubartowski", "rawicki", "polkowicki", _
        "milicki", "piotrkowski", "przysuski", "lubi�ski", "radomski", "pu�awski", "opoczy�ski", "trzebnicki", "boles�awiecki", "ostrzeszowski", "be�chatowski", "wo�owski", "zgorzelecki", "Radom", "zwole�ski", "lubelski", "��czy�ski", "Piotrk�w Trybunalski", "che�mski", "ole�nicki", "wieruszowski", "legnicki", "wielu�ski", "k�pi�ski", "szyd�owiecki", "Lublin", "konecki", "opolski (lubelskie)", "�widnicki (lubelskie)", "lipski", "z�otoryjski", "paj�cza�ski", "radomszcza�ski", "�redzki (dolny�l�sk)", "wroc�awski", "Legnica", "Che�m", "lw�wecki", "luba�ski", "starachowicki", "Wroc�aw", "skar�yski", "namys�owski", "krasnostawski", "kluczborski", "jaworski", "oleski", "opatowski", "o�awski", "ostrowiecki", "kra�nicki", "k�obucki", "hrubieszowski", "kielecki", "jeleniog�rski", "�widnicki (dolno�l�skie)", "cz�stochowski", "w�oszczowski", "brzeski (opolskie)", "zamojski", _
        "Jelenia G�ra", "opolski (opolskie)", "janowski", "bi�gorajski", "Kielce", "kamiennog�rski", "strzeli�ski", "sandomierski", "wa�brzyski", "Cz�stochowa", "j�drzejowski", "lubliniecki", "stalowowolski", "dzier�oniowski", "Zamo��", "tomaszowski (lubelskie)", "tarnobrzeski", "z�bkowicki", "zawiercia�ski", "staszowski", "Opole", "strzelecki", "myszkowski", "k�odzki", "pi�czowski", "Tarnobrzeg", "ni�a�ski", "nyski", "buski", "tarnog�rski", "krapkowicki", "prudnicki", "gliwicki", "b�dzi�ski", "mielecki", "miechowski", "kolbuszowski", "lubaczowski", "D�browa-G�rnicza", "okulski", "le�ajski", "Bytom", "Piekary �l�skie", "k�dzierzy�sko-kozielski", "rzeszowski", "kazimierski", "Gliwice", "Zabrze", "d�browski", "przeworski", "Siemianowice �l�skie", "Chorz�w", "Ruda �l�ska", "proszowicki", "�wi�toch�owice", "Sosnowiec", "g�ubczycki", "krakowski", "Katowice", "�a�cucki", _
        "jaros�awski", "Mys�owice", "raciborski", "miko�owski", "tarnowski", "chrzanowski", "brzeski (ma�opolskie)", "ropczycko-s�dziszowski", "rybnicki", "d�bicki", "Rybnik", "Tychy", "bieru�sko-l�dzi�ski", "boche�ski", "Krak�w", "Rzesz�w", "o�wi�cimski", "wielicki", "pszczy�ski", "Tarn�w", "wodzis�awski", "wadowicki", "przemyski", "strzy�owski", "bielski (�l�skie)", "cieszy�ski", "my�lenicki", "jasielski", "Bielsko-Bia�a", "kro�nie�ski", "brzozowski", "Przemy�l", "limanowski", "�ywiecki", "suski", "nowos�decki", "gorlicki", "Krosno", "bieszczadzki", "sanocki", "Nowy S�cz", "nowotarski", "leski", "tatrza�ski", "�ory", "Jastrz�bie-Zdr�j", "Jaworzno", "��dzki wschodni")) _
        .Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(220, 20, 60)
            .Solid
        End With
    
       With ActiveSheet.Shapes.Range(Array( _
        "bialski", "�redzki (wielkopolskie)", "grodziski (wielkopolskie)", "�owicki", "kro�nie�ski (lubuskie)", "wolszty�ski", "zielonog�rski", "otwocki", "ko�cia�ski", "pruszkowski", "��czycki", "grodziski (mazowieckie)", "�remski", "turecki", "�yrardowski", "piaseczy�ski", "skierniewicki", "jaroci�ski", "Bia�a Podlaska", "�ukowski", "pleszewski", "podd�bicki", "nowosolski", "garwoli�ski", "kaliski", "Zielona G�ra", "wschowski", "zgierski", "radzy�ski", "leszczy�ski", "Skierniewice", "gosty�ski", "gr�jecki", "brzezi�ski", "�arski", "krotoszy�ski", "rawski", "Leszno", "kozienicki", "w�odawski", "��d�", "parczewski", "sieradzki", "rycki", "g�ogowski", "�aga�ski", "ostrowski (wielkopolskie)", "pabianicki", "g�rowski", "Kalisz", "tomaszowski (��dzkie)", "zdu�skowolski", "�aski", "bia�obrzeski", "lubartowski", "rawicki", "polkowicki", _
        "milicki", "piotrkowski", "przysuski", "lubi�ski", "radomski", "pu�awski", "opoczy�ski", "trzebnicki", "boles�awiecki", "ostrzeszowski", "be�chatowski", "wo�owski", "zgorzelecki", "Radom", "zwole�ski", "lubelski", "��czy�ski", "Piotrk�w Trybunalski", "che�mski", "ole�nicki", "wieruszowski", "legnicki", "wielu�ski", "k�pi�ski", "szyd�owiecki", "Lublin", "konecki", "opolski (lubelskie)", "�widnicki (lubelskie)", "lipski", "z�otoryjski", "paj�cza�ski", "radomszcza�ski", "�redzki (dolny�l�sk)", "wroc�awski", "Legnica", "Che�m", "lw�wecki", "luba�ski", "starachowicki", "Wroc�aw", "skar�yski", "namys�owski", "krasnostawski", "kluczborski", "jaworski", "oleski", "opatowski", "o�awski", "ostrowiecki", "kra�nicki", "k�obucki", "hrubieszowski", "kielecki", "jeleniog�rski", "�widnicki (dolno�l�skie)", "cz�stochowski", "w�oszczowski", "brzeski (opolskie)", "zamojski", _
        "Jelenia G�ra", "opolski (opolskie)", "janowski", "bi�gorajski", "Kielce", "kamiennog�rski", "strzeli�ski", "sandomierski", "wa�brzyski", "Cz�stochowa", "j�drzejowski", "lubliniecki", "stalowowolski", "dzier�oniowski", "Zamo��", "tomaszowski (lubelskie)", "tarnobrzeski", "z�bkowicki", "zawiercia�ski", "staszowski", "Opole", "strzelecki", "myszkowski", "k�odzki", "pi�czowski", "Tarnobrzeg", "ni�a�ski", "nyski", "buski", "tarnog�rski", "krapkowicki", "prudnicki", "gliwicki", "b�dzi�ski", "mielecki", "miechowski", "kolbuszowski", "lubaczowski", "D�browa-G�rnicza", "okulski", "le�ajski", "Bytom", "Piekary �l�skie", "k�dzierzy�sko-kozielski", "rzeszowski", "kazimierski", "Gliwice", "Zabrze", "d�browski", "przeworski", "Siemianowice �l�skie", "Chorz�w", "Ruda �l�ska", "proszowicki", "�wi�toch�owice", "Sosnowiec", "g�ubczycki", "krakowski", "Katowice", "�a�cucki", _
        "jaros�awski", "Mys�owice", "raciborski", "miko�owski", "tarnowski", "chrzanowski", "brzeski (ma�opolskie)", "ropczycko-s�dziszowski", "rybnicki", "d�bicki", "Rybnik", "Tychy", "bieru�sko-l�dzi�ski", "boche�ski", "Krak�w", "Rzesz�w", "o�wi�cimski", "wielicki", "pszczy�ski", "Tarn�w", "wodzis�awski", "wadowicki", "przemyski", "strzy�owski", "bielski (�l�skie)", "cieszy�ski", "my�lenicki", "jasielski", "Bielsko-Bia�a", "kro�nie�ski", "brzozowski", "Przemy�l", "limanowski", "�ywiecki", "suski", "nowos�decki", "gorlicki", "Krosno", "bieszczadzki", "sanocki", "Nowy S�cz", "nowotarski", "leski", "tatrza�ski", "�ory", "Jastrz�bie-Zdr�j", "Jaworzno", "��dzki wschodni")) _
        .Line
            .Visible = msoTrue
            .ForeColor.RGB = RGB(220, 20, 60)
            .Visible = msoTrue
            .Weight = 2.5
        End With
        
        ActiveSheet.Shapes.Range(Array( _
        "bialski", "�redzki (wielkopolskie)", "grodziski (wielkopolskie)", "�owicki", "kro�nie�ski (lubuskie)", "wolszty�ski", "zielonog�rski", "otwocki", "ko�cia�ski", "pruszkowski", "��czycki", "grodziski (mazowieckie)", "�remski", "turecki", "�yrardowski", "piaseczy�ski", "skierniewicki", "jaroci�ski", "Bia�a Podlaska", "�ukowski", "pleszewski", "podd�bicki", "nowosolski", "garwoli�ski", "kaliski", "Zielona G�ra", "wschowski", "zgierski", "radzy�ski", "leszczy�ski", "Skierniewice", "gosty�ski", "gr�jecki", "brzezi�ski", "�arski", "krotoszy�ski", "rawski", "Leszno", "kozienicki", "w�odawski", "��d�", "parczewski", "sieradzki", "rycki", "g�ogowski", "�aga�ski", "ostrowski (wielkopolskie)", "pabianicki", "g�rowski", "Kalisz", "tomaszowski (��dzkie)", "zdu�skowolski", "�aski", "bia�obrzeski", "lubartowski", "rawicki", "polkowicki", _
        "milicki", "piotrkowski", "przysuski", "lubi�ski", "radomski", "pu�awski", "opoczy�ski", "trzebnicki", "boles�awiecki", "ostrzeszowski", "be�chatowski", "wo�owski", "zgorzelecki", "Radom", "zwole�ski", "lubelski", "��czy�ski", "Piotrk�w Trybunalski", "che�mski", "ole�nicki", "wieruszowski", "legnicki", "wielu�ski", "k�pi�ski", "szyd�owiecki", "Lublin", "konecki", "opolski (lubelskie)", "�widnicki (lubelskie)", "lipski", "z�otoryjski", "paj�cza�ski", "radomszcza�ski", "�redzki (dolny�l�sk)", "wroc�awski", "Legnica", "Che�m", "lw�wecki", "luba�ski", "starachowicki", "Wroc�aw", "skar�yski", "namys�owski", "krasnostawski", "kluczborski", "jaworski", "oleski", "opatowski", "o�awski", "ostrowiecki", "kra�nicki", "k�obucki", "hrubieszowski", "kielecki", "jeleniog�rski", "�widnicki (dolno�l�skie)", "cz�stochowski", "w�oszczowski", "brzeski (opolskie)", "zamojski", _
        "Jelenia G�ra", "opolski (opolskie)", "janowski", "bi�gorajski", "Kielce", "kamiennog�rski", "strzeli�ski", "sandomierski", "wa�brzyski", "Cz�stochowa", "j�drzejowski", "lubliniecki", "stalowowolski", "dzier�oniowski", "Zamo��", "tomaszowski (lubelskie)", "tarnobrzeski", "z�bkowicki", "zawiercia�ski", "staszowski", "Opole", "strzelecki", "myszkowski", "k�odzki", "pi�czowski", "Tarnobrzeg", "ni�a�ski", "nyski", "buski", "tarnog�rski", "krapkowicki", "prudnicki", "gliwicki", "b�dzi�ski", "mielecki", "miechowski", "kolbuszowski", "lubaczowski", "D�browa-G�rnicza", "okulski", "le�ajski", "Bytom", "Piekary �l�skie", "k�dzierzy�sko-kozielski", "rzeszowski", "kazimierski", "Gliwice", "Zabrze", "d�browski", "przeworski", "Siemianowice �l�skie", "Chorz�w", "Ruda �l�ska", "proszowicki", "�wi�toch�owice", "Sosnowiec", "g�ubczycki", "krakowski", "Katowice", "�a�cucki", _
        "jaros�awski", "Mys�owice", "raciborski", "miko�owski", "tarnowski", "chrzanowski", "brzeski (ma�opolskie)", "ropczycko-s�dziszowski", "rybnicki", "d�bicki", "Rybnik", "Tychy", "bieru�sko-l�dzi�ski", "boche�ski", "Krak�w", "Rzesz�w", "o�wi�cimski", "wielicki", "pszczy�ski", "Tarn�w", "wodzis�awski", "wadowicki", "przemyski", "strzy�owski", "bielski (�l�skie)", "cieszy�ski", "my�lenicki", "jasielski", "Bielsko-Bia�a", "kro�nie�ski", "brzozowski", "Przemy�l", "limanowski", "�ywiecki", "suski", "nowos�decki", "gorlicki", "Krosno", "bieszczadzki", "sanocki", "Nowy S�cz", "nowotarski", "leski", "tatrza�ski", "�ory", "Jastrz�bie-Zdr�j", "Jaworzno", "��dzki wschodni")) _
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
    
            .ChartTitle.Text = "Polska - prognoza liczby ludno�ci na lata 2011-2035"
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
       
        With ActiveSheet.Shapes.Range(Array("pucki", "wejherowski", "l�borski", "s�upski" _
            , "bytowski", "cz�uchowski", "chojnicki", "ko�cierski", "kartuski", "gda�ski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzy�ski")).Fill
            
            .ForeColor.RGB = RGB(255, 192, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("pucki", "wejherowski", "l�borski", "s�upski" _
            , "bytowski", "cz�uchowski", "chojnicki", "ko�cierski", "kartuski", "gda�ski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzy�ski")).Line
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("s�awie�ski", "bia�ogardzki", "koszali�ski", _
            "szczecinecki", "wa�ecki", "drawski", "�widwi�ski", "ko�obrzeski", "gryficki", _
            "�obeski", "�winouj�cie", "goleniowski", "stargardzki", "choszcze�ski", _
            "pyrzycki", "my�liborski", "gryfi�ski", "policki", "kamie�ski")).Fill
            
            .ForeColor.RGB = RGB(192, 0, 0)
        End With

        With ActiveSheet.Shapes.Range(Array("s�awie�ski", "bia�ogardzki", "koszali�ski", _
            "szczecinecki", "wa�ecki", "drawski", "�widwi�ski", "ko�obrzeski", "gryficki", _
            "�obeski", "�winouj�cie", "goleniowski", "stargardzki", "choszcze�ski", _
            "pyrzycki", "my�liborski", "gryfi�ski", "policki", "kamie�ski")).Line
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("braniewski", "elbl�ski", "i�awski", _
            "nowomiejski", "dzia�dowski", "ostr�dzki", "lidzbarski", "bartoszycki", _
            "olszty�ski", "nidzicki", "szczycie�ski", "mr�gowski", "k�trzy�ski", _
            "w�gorzewski", "go�dapski", "olecki", "gi�ycki", "e�cki", "piski")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent3
        End With
        
        With ActiveSheet.Shapes.Range(Array("braniewski", "elbl�ski", "i�awski", _
            "nowomiejski", "dzia�dowski", "ostr�dzki", "lidzbarski", "bartoszycki", _
            "olszty�ski", "nidzicki", "szczycie�ski", "mr�gowski", "k�trzy�ski", _
            "w�gorzewski", "go�dapski", "olecki", "gi�ycki", "e�cki", "piski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("s�pole�ski", "tucholski", "che�mi�ski", _
            "�wiecki", "grudzi�dzki", "w�brzeski", "brodnicki", "rypi�ski", _
            "golubsko-dobrzy�ski", "lipnowski", "w�oc�awski", "radziejowski", _
            "aleksandrowski", "toru�ski", "inowroc�awski", "mogile�ski", "bydgoski", _
            "�ni�ski")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent4
        End With
        
        With ActiveSheet.Shapes.Range(Array("s�pole�ski", "tucholski", "che�mi�ski", _
            "�wiecki", "grudzi�dzki", "w�brzeski", "brodnicki", "rypi�ski", _
            "golubsko-dobrzy�ski", "lipnowski", "w�oc�awski", "radziejowski", _
            "aleksandrowski", "toru�ski", "inowroc�awski", "mogile�ski", "bydgoski", _
            "�ni�ski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("z�otowski", "pilski", "nakielski", _
            "chodzieski", "w�growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "mi�dzychodzki", "nowotomyski", "wolszty�ski", "grodziski (wielkopolskie)", _
            "pozna�ski", "gnie�nie�ski", "kolski", "koni�ski", "s�upecki", "wrzesi�ski", _
            "�remski", "jaroci�ski", "�redzki (wielkopolskie)", "leszczy�ski", "gosty�ski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszy�ski", "ostrzeszowski", "k�pi�ski", "turecki", "ko�cia�ski")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent5
        End With
        
        With ActiveSheet.Shapes.Range(Array("z�otowski", "pilski", "nakielski", _
            "chodzieski", "w�growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "mi�dzychodzki", "nowotomyski", "wolszty�ski", "grodziski (wielkopolskie)", _
            "pozna�ski", "gnie�nie�ski", "kolski", "koni�ski", "s�upecki", "wrzesi�ski", _
            "�remski", "jaroci�ski", "�redzki (wielkopolskie)", "leszczy�ski", "gosty�ski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszy�ski", "ostrzeszowski", "k�pi�ski", "turecki", "ko�cia�ski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
     
        With ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "gorzowski", _
            "sul�ci�ski", "s�ubicki", "mi�dzyrzecki", "�wiebodzi�ski", "zielonog�rski", _
            "kro�nie�ski (lubuskie)", "�arski", "nowosolski", "�aga�ski", "wschowski")). _
            Fill
            
            .ForeColor.RGB = RGB(165, 27, 96)
        End With
        
        With ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "gorzowski", _
            "sul�ci�ski", "s�ubicki", "mi�dzyrzecki", "�wiebodzi�ski", "zielonog�rski", _
            "kro�nie�ski (lubuskie)", "�arski", "nowosolski", "�aga�ski", "wschowski")). _
            Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("wa�brzyski", "kamiennog�rski", "z�otoryjski" _
            , "luba�ski", "zgorzelecki", "boles�awiecki", "polkowicki", "g�ogowski", _
            "g�rowski", "wo�owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "ole�nicki", "wroc�awski", "o�awski", "strzeli�ski", "dzier�oniowski", _
            "z�bkowicki", "lubi�ski", "�redzki (dolny�l�sk)", "�widnicki (dolno�l�skie)", _
            "lw�wecki", "jeleniog�rski", "k�odzki")).Fill
            
            .ForeColor.RGB = RGB(255, 0, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("wa�brzyski", "kamiennog�rski", "z�otoryjski" _
            , "luba�ski", "zgorzelecki", "boles�awiecki", "polkowicki", "g�ogowski", _
            "g�rowski", "wo�owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "ole�nicki", "wroc�awski", "o�awski", "strzeli�ski", "dzier�oniowski", _
            "z�bkowicki", "lubi�ski", "�redzki (dolny�l�sk)", "�widnicki (dolno�l�skie)", _
            "lw�wecki", "jeleniog�rski", "k�odzki")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("suwalski", "sejne�ski", "augustowski", _
            "grajewski", "kolne�ski", "moniecki", "sok�lski", "bia�ostocki", "�om�y�ski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        End With
        
        With ActiveSheet.Shapes.Range(Array("suwalski", "sejne�ski", "augustowski", _
            "grajewski", "kolne�ski", "moniecki", "sok�lski", "bia�ostocki", "�om�y�ski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("ostro��cki", "przasnyski", "ciechanowski", _
            "m�awski", "�uromi�ski", "sierpecki", "gostyni�ski", "p�ocki", "p�o�ski", _
            "pu�tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko�owski", "�osicki", "siedlecki", "w�growski", _
            "wo�omi�ski", "mi�ski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "�yrardowski", "piaseczy�ski", "otwocki", "gr�jecki", "przysuski", _
            "bia�obrzeski", "kozienicki", "garwoli�ski", "zwole�ski", "szyd�owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Fill
            
            .ForeColor.RGB = RGB(255, 204, 51)
        End With
        
        With ActiveSheet.Shapes.Range(Array("ostro��cki", "przasnyski", "ciechanowski", _
            "m�awski", "�uromi�ski", "sierpecki", "gostyni�ski", "p�ocki", "p�o�ski", _
            "pu�tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko�owski", "�osicki", "siedlecki", "w�growski", _
            "wo�omi�ski", "mi�ski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "�yrardowski", "piaseczy�ski", "otwocki", "gr�jecki", "przysuski", _
            "bia�obrzeski", "kozienicki", "garwoli�ski", "zwole�ski", "szyd�owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("rawski", "skierniewicki", "�owicki", _
            "kutnowski", "��czycki", "zgierski", "brzezi�ski", "tomaszowski (��dzkie)", _
            "opoczy�ski", "��dzki wschodni", "piotrkowski", "radomszcza�ski", "paj�cza�ski" _
            , "be�chatowski", "pabianicki", "�aski", "zdu�skowolski", "podd�bicki", _
            "sieradzki", "wielu�ski", "wieruszowski")).Fill
            
            .ForeColor.RGB = RGB(0, 176, 80)
        End With
        
        With ActiveSheet.Shapes.Range(Array("rawski", "skierniewicki", "�owicki", _
            "kutnowski", "��czycki", "zgierski", "brzezi�ski", "tomaszowski (��dzkie)", _
            "opoczy�ski", "��dzki wschodni", "piotrkowski", "radomszcza�ski", "paj�cza�ski" _
            , "be�chatowski", "pabianicki", "�aski", "zdu�skowolski", "podd�bicki", _
            "sieradzki", "wielu�ski", "wieruszowski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("g�ubczycki", "k�dzierzy�sko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys�owski", "brzeski (opolskie)", "nyski")).Fill
            
            .ForeColor.RGB = RGB(0, 0, 204)
        End With
        
        With ActiveSheet.Shapes.Range(Array("g�ubczycki", "k�dzierzy�sko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys�owski", "brzeski (opolskie)", "nyski")).Line
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("bieru�sko-l�dzi�ski", "pszczy�ski", _
            "bielski (�l�skie)", "cieszy�ski", "wodzis�awski", "rybnicki", "raciborski", _
            "gliwicki", "miko�owski", "b�dzi�ski", "tarnog�rski", "zawiercia�ski", _
            "myszkowski", "cz�stochowski", "lubliniecki", "k�obucki", "�ywiecki")).Fill
            
            .ForeColor.RGB = RGB(204, 153, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("bieru�sko-l�dzi�ski", "pszczy�ski", _
            "bielski (�l�skie)", "cieszy�ski", "wodzis�awski", "rybnicki", "raciborski", _
            "gliwicki", "miko�owski", "b�dzi�ski", "tarnog�rski", "zawiercia�ski", _
            "myszkowski", "cz�stochowski", "lubliniecki", "k�obucki", "�ywiecki")).Line
            
             .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("konecki", "skar�yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w�oszczowski", "j�drzejowski", "pi�czowski", "buski", "kazimierski")).Fill
            
        .ForeColor.RGB = RGB(204, 51, 255)

        End With
        
        With ActiveSheet.Shapes.Range(Array("konecki", "skar�yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w�oszczowski", "j�drzejowski", "pi�czowski", "buski", "kazimierski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("bialski", "radzy�ski", "�ukowski", "rycki", _
            "parczewski", "w�odawski", "lubartowski", "pu�awski", "opolski (lubelskie)", _
            "lubelski", "kra�nicki", "�widnicki (lubelskie)", "che�mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi�gorajski", "��czy�ski", _
            "tomaszowski (lubelskie)")).Fill
            
            .ForeColor.RGB = RGB(102, 0, 204)
        End With
        
        With ActiveSheet.Shapes.Range(Array("bialski", "radzy�ski", "�ukowski", "rycki", _
            "parczewski", "w�odawski", "lubartowski", "pu�awski", "opolski (lubelskie)", _
            "lubelski", "kra�nicki", "�widnicki (lubelskie)", "che�mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi�gorajski", "��czy�ski", _
            "tomaszowski (lubelskie)")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("jaros�awski", "przemyski", "le�ajski", _
            "przeworski", "�a�cucki", "ni�a�ski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-s�dziszowski", "mielecki", "d�bicki", _
            "brzozowski", "kro�nie�ski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy�owski", "lubaczowski")).Fill
            
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 64, 64)
        End With
        
        With ActiveSheet.Shapes.Range(Array("jaros�awski", "przemyski", "le�ajski", _
            "przeworski", "�a�cucki", "ni�a�ski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-s�dziszowski", "mielecki", "d�bicki", _
            "brzozowski", "kro�nie�ski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy�owski", "lubaczowski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("d�browski", "gorlicki", "nowos�decki", _
            "boche�ski", "brzeski (ma�opolskie)", "tarnowski", "limanowski", "wielicki", _
            "my�lenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "o�wi�cimski", "chrzanowski", "okulski", "nowotarski", "tatrza�ski")).Fill
            
            .ForeColor.RGB = RGB(204, 51, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("d�browski", "gorlicki", "nowos�decki", _
            "boche�ski", "brzeski (ma�opolskie)", "tarnowski", "limanowski", "wielicki", _
            "my�lenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "o�wi�cimski", "chrzanowski", "okulski", "nowotarski", "tatrza�ski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("Gdynia", "Sopot", "Elbl�g", "S�upsk", _
            "Koszalin", "Grudzi�dz", "�winouj�cie", "Suwa�ki", "�om�a", "Ostro��ka", "P�ock", "W�oc�awek" _
            , "Konin", "Siedlce", "Bia�a Podlaska", "Skierniewice", "Piotrk�w Trybunalski" _
            , "Kalisz", "Leszno", "Legnica", "Jelenia G�ra", "Radom", "Che�m", "Zamo��", _
            "Tarnobrzeg", "Tarn�w", "Nowy S�cz", "Krosno", "Przemy�l", "Sosnowiec", _
            "D�browa-G�rnicza", "Jaworzno", "Bielsko-Bia�a", "Mys�owice", "Tychy", "�ory", _
            "Jastrz�bie-Zdr�j", "Rybnik", "Ruda �l�ska", "Bytom", "Piekary �l�skie", _
            "Zabrze", "Gliwice", "Cz�stochowa", "�wi�toch�owice")).Fill
            
            .ForeColor.RGB = RGB(255, 69, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Gdynia", "Sopot", "Elbl�g", "S�upsk", _
            "Koszalin", "�winouj�cie", "Grudzi�dz", "Suwa�ki", "�om�a", "Ostro��ka", "P�ock", "W�oc�awek" _
            , "Konin", "Siedlce", "Bia�a Podlaska", "Skierniewice", "Piotrk�w Trybunalski" _
            , "Kalisz", "Leszno", "Legnica", "Jelenia G�ra", "Radom", "Che�m", "Zamo��", _
            "Tarnobrzeg", "Tarn�w", "Nowy S�cz", "Krosno", "Przemy�l", "Sosnowiec", _
            "D�browa-G�rnicza", "Jaworzno", "Bielsko-Bia�a", "Mys�owice", "Tychy", "�ory", _
            "Jastrz�bie-Zdr�j", "Rybnik", "Ruda �l�ska", "Bytom", "Piekary �l�skie", _
            "Zabrze", "Gliwice", "Cz�stochowa", "�wi�toch�owice")).Line
            
            .ForeColor.RGB = RGB(0, 0, 0)
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("Gda�sk", "Szczecin", "Gorz�w Wielkopolski", _
            "Pozna�", "Bydgoszcz", "Toru�", "Olsztyn", "Bia�ystok", "Warszawa", "��d�", _
            "Zielona G�ra", "Wroc�aw", "Opole", "Kielce", "Lublin", "Rzesz�w", "Katowice", _
            "Krak�w")).Fill
            
            .ForeColor.RGB = RGB(255, 102, 153)
         End With
         
        With ActiveSheet.Shapes.Range(Array("Gda�sk", "Szczecin", "Gorz�w Wielkopolski", _
            "Pozna�", "Bydgoszcz", "Toru�", "Olsztyn", "Bia�ystok", "Warszawa", "��d�", _
            "Zielona G�ra", "Wroc�aw", "Opole", "Kielce", "Lublin", "Rzesz�w", "Katowice", _
            "Krak�w")).Line
            
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
    
            .ChartTitle.Text = "Miasto na prawach powiatu Warszawa - prognoza liczby ludno�ci na lata 2011-2035"
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
        ActiveSheet.Shapes.Range(Array("warmi�sko-mazurskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("kujawsko-pomorskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("wielkopolskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("lubuskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("dolno�l�skie")).Ungroup
        ActiveSheet.Shapes.Range(Array("opolskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("�l�skie")).Ungroup
        ActiveSheet.Shapes.Range(Array("mazowieckie")).Ungroup
        ActiveSheet.Shapes.Range(Array("podlaskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("�wi�tokrzyskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("lubelskie")).Ungroup
        ActiveSheet.Shapes.Range(Array("��dzkie")).Ungroup
        ActiveSheet.Shapes.Range(Array("podkarpackie")).Ungroup
        ActiveSheet.Shapes.Range(Array("ma�opolskie")).Ungroup
    
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
       
        With ActiveSheet.Shapes.Range(Array("pucki", "wejherowski", "l�borski", "s�upski" _
            , "bytowski", "cz�uchowski", "chojnicki", "ko�cierski", "kartuski", "gda�ski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzy�ski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("s�awie�ski", "bia�ogardzki", "koszali�ski", _
            "szczecinecki", "wa�ecki", "drawski", "�widwi�ski", "ko�obrzeski", "gryficki", _
            "�obeski", "�winouj�cie", "goleniowski", "stargardzki", "choszcze�ski", _
            "pyrzycki", "my�liborski", "gryfi�ski", "policki", "kamie�ski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("braniewski", "elbl�ski", "i�awski", _
            "nowomiejski", "dzia�dowski", "ostr�dzki", "lidzbarski", "bartoszycki", _
            "olszty�ski", "nidzicki", "szczycie�ski", "mr�gowski", "k�trzy�ski", _
            "w�gorzewski", "go�dapski", "olecki", "gi�ycki", "e�cki", "piski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("s�pole�ski", "tucholski", "che�mi�ski", _
            "�wiecki", "grudzi�dzki", "w�brzeski", "brodnicki", "rypi�ski", _
            "golubsko-dobrzy�ski", "lipnowski", "w�oc�awski", "radziejowski", _
            "aleksandrowski", "toru�ski", "inowroc�awski", "mogile�ski", "bydgoski", _
            "�ni�ski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("z�otowski", "pilski", "nakielski", _
            "chodzieski", "w�growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "mi�dzychodzki", "nowotomyski", "wolszty�ski", "grodziski (wielkopolskie)", _
            "pozna�ski", "gnie�nie�ski", "kolski", "koni�ski", "s�upecki", "wrzesi�ski", _
            "�remski", "jaroci�ski", "�redzki (wielkopolskie)", "leszczy�ski", "gosty�ski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszy�ski", "ostrzeszowski", "k�pi�ski", "turecki", "ko�cia�ski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
     
        With ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "gorzowski", _
            "sul�ci�ski", "s�ubicki", "mi�dzyrzecki", "�wiebodzi�ski", "zielonog�rski", _
            "kro�nie�ski (lubuskie)", "�arski", "nowosolski", "�aga�ski", "wschowski")). _
            Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("wa�brzyski", "kamiennog�rski", "z�otoryjski" _
            , "luba�ski", "zgorzelecki", "boles�awiecki", "polkowicki", "g�ogowski", _
            "g�rowski", "wo�owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "ole�nicki", "wroc�awski", "o�awski", "strzeli�ski", "dzier�oniowski", _
            "z�bkowicki", "lubi�ski", "�redzki (dolny�l�sk)", "�widnicki (dolno�l�skie)", _
            "lw�wecki", "jeleniog�rski", "k�odzki")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("suwalski", "sejne�ski", "augustowski", _
            "grajewski", "kolne�ski", "moniecki", "sok�lski", "bia�ostocki", "�om�y�ski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("ostro��cki", "przasnyski", "ciechanowski", _
            "m�awski", "�uromi�ski", "sierpecki", "gostyni�ski", "p�ocki", "p�o�ski", _
            "pu�tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko�owski", "�osicki", "siedlecki", "w�growski", _
            "wo�omi�ski", "mi�ski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "�yrardowski", "piaseczy�ski", "otwocki", "gr�jecki", "przysuski", _
            "bia�obrzeski", "kozienicki", "garwoli�ski", "zwole�ski", "szyd�owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("rawski", "skierniewicki", "�owicki", _
            "kutnowski", "��czycki", "zgierski", "brzezi�ski", "tomaszowski (��dzkie)", _
            "opoczy�ski", "��dzki wschodni", "piotrkowski", "radomszcza�ski", "paj�cza�ski" _
            , "be�chatowski", "pabianicki", "�aski", "zdu�skowolski", "podd�bicki", _
            "sieradzki", "wielu�ski", "wieruszowski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("g�ubczycki", "k�dzierzy�sko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys�owski", "brzeski (opolskie)", "nyski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("bieru�sko-l�dzi�ski", "pszczy�ski", _
            "bielski (�l�skie)", "cieszy�ski", "wodzis�awski", "rybnicki", "raciborski", _
            "gliwicki", "miko�owski", "b�dzi�ski", "tarnog�rski", "zawiercia�ski", _
            "myszkowski", "cz�stochowski", "lubliniecki", "k�obucki", "�ywiecki")).Line
            
             .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("konecki", "skar�yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w�oszczowski", "j�drzejowski", "pi�czowski", "buski", "kazimierski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("bialski", "radzy�ski", "�ukowski", "rycki", _
            "parczewski", "w�odawski", "lubartowski", "pu�awski", "opolski (lubelskie)", _
            "lubelski", "kra�nicki", "�widnicki (lubelskie)", "che�mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi�gorajski", "��czy�ski", _
            "tomaszowski (lubelskie)")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("jaros�awski", "przemyski", "le�ajski", _
            "przeworski", "�a�cucki", "ni�a�ski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-s�dziszowski", "mielecki", "d�bicki", _
            "brzozowski", "kro�nie�ski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy�owski", "lubaczowski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("d�browski", "gorlicki", "nowos�decki", _
            "boche�ski", "brzeski (ma�opolskie)", "tarnowski", "limanowski", "wielicki", _
            "my�lenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "o�wi�cimski", "chrzanowski", "okulski", "nowotarski", "tatrza�ski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorText1
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("Gdynia", "Sopot", "Elbl�g", "S�upsk", _
            "Koszalin", "�winouj�cie", "Grudzi�dz", "Suwa�ki", "�om�a", "Ostro��ka", "P�ock", "W�oc�awek" _
            , "Konin", "Siedlce", "Bia�a Podlaska", "Skierniewice", "Piotrk�w Trybunalski" _
            , "Kalisz", "Leszno", "Legnica", "Jelenia G�ra", "Radom", "Che�m", "Zamo��", _
            "Tarnobrzeg", "Tarn�w", "Nowy S�cz", "Krosno", "Przemy�l", "Sosnowiec", _
            "D�browa-G�rnicza", "Jaworzno", "Bielsko-Bia�a", "Mys�owice", "Tychy", "�ory", _
            "Jastrz�bie-Zdr�j", "Rybnik", "Ruda �l�ska", "Bytom", "Piekary �l�skie", _
            "Zabrze", "Gliwice", "Cz�stochowa", "�wi�toch�owice")).Fill
            
            .ForeColor.RGB = RGB(255, 69, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Gdynia", "Sopot", "Elbl�g", "S�upsk", _
            "Koszalin", "�winouj�cie", "Grudzi�dz", "Suwa�ki", "�om�a", "Ostro��ka", "P�ock", "W�oc�awek" _
            , "Konin", "Siedlce", "Bia�a Podlaska", "Skierniewice", "Piotrk�w Trybunalski" _
            , "Kalisz", "Leszno", "Legnica", "Jelenia G�ra", "Radom", "Che�m", "Zamo��", _
            "Tarnobrzeg", "Tarn�w", "Nowy S�cz", "Krosno", "Przemy�l", "Sosnowiec", _
            "D�browa-G�rnicza", "Jaworzno", "Bielsko-Bia�a", "Mys�owice", "Tychy", "�ory", _
            "Jastrz�bie-Zdr�j", "Rybnik", "Ruda �l�ska", "Bytom", "Piekary �l�skie", _
            "Zabrze", "Gliwice", "Cz�stochowa", "�wi�toch�owice")).Line
            
            .ForeColor.RGB = RGB(0, 0, 0)
            .Weight = 1
        End With
    
        With ActiveSheet.Shapes.Range(Array("Gda�sk", "Szczecin", "Gorz�w Wielkopolski", _
            "Pozna�", "Bydgoszcz", "Toru�", "Olsztyn", "Bia�ystok", "Warszawa", "��d�", _
            "Zielona G�ra", "Wroc�aw", "Opole", "Kielce", "Lublin", "Rzesz�w", "Katowice", _
            "Krak�w")).Fill
            
            .ForeColor.RGB = RGB(255, 102, 153)
         End With
         
        With ActiveSheet.Shapes.Range(Array("Gda�sk", "Szczecin", "Gorz�w Wielkopolski", _
            "Pozna�", "Bydgoszcz", "Toru�", "Olsztyn", "Bia�ystok", "Warszawa", "��d�", _
            "Zielona G�ra", "Wroc�aw", "Opole", "Kielce", "Lublin", "Rzesz�w", "Katowice", _
            "Krak�w")).Line
            
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
    
            .ChartTitle.Text = "Miasto na prawach powiatu Warszawa - prognoza liczby ludno�ci na lata 2011-2035"
            .SeriesCollection(1).Values = Worksheets("Mapka dane").Range("C" & 383 & "," & "F" & 383 & "," & "I" & 383 & "," & "L" & 383 & "," & "O" & 383 & "," & "R" & 383 & "," & "U" & 383 & "," & "X" & 383 & "," & "AA" & 383 & "," & "AD" & 383 & "," & "AG" & 383 & "," & "AJ" & 383 & "," & "AM" & 383 & "," & "AP" & 383 & "," & "AS" & 383 & "," & "AV" & 383 & "," & "AY" & 383 & "," & "BB" & 383 & "," & "BE" & 383 & "," & "BK" & 383 & "," & "BN" & 383 & "," & "BQ" & 383 & "," & "BT" & 383 & "," & "BW" & 383)
            .SeriesCollection(2).Values = Worksheets("Mapka dane").Range("D" & 383 & "," & "G" & 383 & "," & "J" & 383 & "," & "M" & 383 & "," & "P" & 383 & "," & "S" & 383 & "," & "V" & 383 & "," & "Y" & 383 & "," & "AB" & 383 & "," & "AE" & 383 & "," & "AH" & 383 & "," & "AK" & 383 & "," & "AN" & 383 & "," & "AQ" & 383 & "," & "AT" & 383 & "," & "AW" & 383 & "," & "AZ" & 383 & "," & "BC" & 383 & "," & "BF" & 383 & "," & "BL" & 383 & "," & "BO" & 383 & "," & "BR" & 383 & "," & "BU" & 383 & "," & "BX" & 383)
            .SeriesCollection(3).Values = Worksheets("Mapka dane").Range("E" & 383 & "," & "H" & 383 & "," & "K" & 383 & "," & "N" & 383 & "," & "Q" & 383 & "," & "T" & 383 & "," & "W" & 383 & "," & "Z" & 383 & "," & "AC" & 383 & "," & "AF" & 383 & "," & "AI" & 383 & "," & "AL" & 383 & "," & "AO" & 383 & "," & "AR" & 383 & "," & "AU" & 383 & "," & "AX" & 383 & "," & "BA" & 383 & "," & "BD" & 383 & "," & "BG" & 383 & "," & "BM" & 383 & "," & "BP" & 383 & "," & "BS" & 383 & "," & "BV" & 383 & "," & "BY" & 383)
    
        End With
          
    End If
    
    RefreshCharts  'odswieza wykresy, �eby update'owa�y si� w czasie dzia�ania excela
     
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
              
        ActiveSheet.Shapes.Range(Array("pucki", "Gdynia", "Sopot", "Gda�sk", "S�upsk", "wejherowski", "l�borski", "s�upski" _
            , "bytowski", "cz�uchowski", "chojnicki", "ko�cierski", "kartuski", "gda�ski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzy�ski")).Group.Name = "pomorskie"
            
        With ActiveSheet.Shapes.Range(Array("pucki", "Gdynia", "Sopot", "Gda�sk", "S�upsk", "wejherowski", "l�borski", "s�upski" _
            , "bytowski", "cz�uchowski", "chojnicki", "ko�cierski", "kartuski", "gda�ski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzy�ski")).Fill
            
            .ForeColor.RGB = RGB(255, 192, 0)
            .Solid
        End With
        
        With ActiveSheet.Shapes.Range(Array("pucki", "Gdynia", "Sopot", "Gda�sk", "S�upsk", "wejherowski", "l�borski", "s�upski" _
            , "bytowski", "cz�uchowski", "chojnicki", "ko�cierski", "kartuski", "gda�ski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzy�ski")).Line
            
            .ForeColor.RGB = RGB(255, 192, 0)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("s�awie�ski", "Szczecin", "Koszalin", "bia�ogardzki", "koszali�ski", _
            "szczecinecki", "wa�ecki", "drawski", "�widwi�ski", "ko�obrzeski", "gryficki", _
            "�obeski", "�winouj�cie", "goleniowski", "stargardzki", "choszcze�ski", _
            "pyrzycki", "my�liborski", "gryfi�ski", "policki", "kamie�ski")).Group.Name = "zachodniopomorskie"
    
        With ActiveSheet.Shapes.Range(Array("s�awie�ski", "Szczecin", "Koszalin", "bia�ogardzki", "koszali�ski", _
            "szczecinecki", "wa�ecki", "drawski", "�widwi�ski", "ko�obrzeski", "gryficki", _
            "�obeski", "�winouj�cie", "goleniowski", "stargardzki", "choszcze�ski", _
            "pyrzycki", "my�liborski", "gryfi�ski", "policki", "kamie�ski")).Fill
            
            .ForeColor.RGB = RGB(192, 0, 0)
        End With

        With ActiveSheet.Shapes.Range(Array("s�awie�ski", "Szczecin", "Koszalin", "bia�ogardzki", "koszali�ski", _
            "szczecinecki", "wa�ecki", "drawski", "�widwi�ski", "ko�obrzeski", "gryficki", _
            "�obeski", "�winouj�cie", "goleniowski", "stargardzki", "choszcze�ski", _
            "pyrzycki", "my�liborski", "gryfi�ski", "policki", "kamie�ski")).Line
            
            .ForeColor.RGB = RGB(192, 0, 0)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Elbl�g", "Olsztyn", "braniewski", "elbl�ski", "i�awski", _
            "nowomiejski", "dzia�dowski", "ostr�dzki", "lidzbarski", "bartoszycki", _
            "olszty�ski", "nidzicki", "szczycie�ski", "mr�gowski", "k�trzy�ski", _
            "w�gorzewski", "go�dapski", "olecki", "gi�ycki", "e�cki", "piski")).Group.Name = "warmi�sko-mazurskie"
    
        With ActiveSheet.Shapes.Range(Array("Elbl�g", "Olsztyn", "braniewski", "elbl�ski", "i�awski", _
            "nowomiejski", "dzia�dowski", "ostr�dzki", "lidzbarski", "bartoszycki", _
            "olszty�ski", "nidzicki", "szczycie�ski", "mr�gowski", "k�trzy�ski", _
            "w�gorzewski", "go�dapski", "olecki", "gi�ycki", "e�cki", "piski")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent3
        End With
        
        With ActiveSheet.Shapes.Range(Array("Elbl�g", "Olsztyn", "braniewski", "elbl�ski", "i�awski", _
            "nowomiejski", "dzia�dowski", "ostr�dzki", "lidzbarski", "bartoszycki", _
            "olszty�ski", "nidzicki", "szczycie�ski", "mr�gowski", "k�trzy�ski", _
            "w�gorzewski", "go�dapski", "olecki", "gi�ycki", "e�cki", "piski")).Line
            .ForeColor.ObjectThemeColor = msoThemeColorAccent3
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("s�pole�ski", "tucholski", "Bydgoszcz", "Toru�", "Grudzi�dz", "W�oc�awek", "che�mi�ski", _
            "�wiecki", "grudzi�dzki", "w�brzeski", "brodnicki", "rypi�ski", _
            "golubsko-dobrzy�ski", "lipnowski", "w�oc�awski", "radziejowski", _
            "aleksandrowski", "toru�ski", "inowroc�awski", "mogile�ski", "bydgoski", _
            "�ni�ski")).Group.Name = "kujawsko-pomorskie"
            
        With ActiveSheet.Shapes.Range(Array("s�pole�ski", "tucholski", "Bydgoszcz", "Toru�", "Grudzi�dz", "W�oc�awek", "che�mi�ski", _
            "�wiecki", "grudzi�dzki", "w�brzeski", "brodnicki", "rypi�ski", _
            "golubsko-dobrzy�ski", "lipnowski", "w�oc�awski", "radziejowski", _
            "aleksandrowski", "toru�ski", "inowroc�awski", "mogile�ski", "bydgoski", _
            "�ni�ski")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent4
        End With
        
        With ActiveSheet.Shapes.Range(Array("s�pole�ski", "tucholski", "Bydgoszcz", "Toru�", "Grudzi�dz", "W�oc�awek", "che�mi�ski", _
            "�wiecki", "grudzi�dzki", "w�brzeski", "brodnicki", "rypi�ski", _
            "golubsko-dobrzy�ski", "lipnowski", "w�oc�awski", "radziejowski", _
            "aleksandrowski", "toru�ski", "inowroc�awski", "mogile�ski", "bydgoski", _
            "�ni�ski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent4
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Pozna�", "Konin", "Kalisz", "Leszno", "z�otowski", "pilski", "nakielski", _
            "chodzieski", "w�growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "mi�dzychodzki", "nowotomyski", "wolszty�ski", "grodziski (wielkopolskie)", _
            "pozna�ski", "gnie�nie�ski", "kolski", "koni�ski", "s�upecki", "wrzesi�ski", _
            "�remski", "jaroci�ski", "�redzki (wielkopolskie)", "leszczy�ski", "gosty�ski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszy�ski", "ostrzeszowski", "k�pi�ski", "turecki", "ko�cia�ski")).Group.Name = "wielkopolskie"
    
        With ActiveSheet.Shapes.Range(Array("Pozna�", "Konin", "Kalisz", "Leszno", "z�otowski", "pilski", "nakielski", _
            "chodzieski", "w�growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "mi�dzychodzki", "nowotomyski", "wolszty�ski", "grodziski (wielkopolskie)", _
            "pozna�ski", "gnie�nie�ski", "kolski", "koni�ski", "s�upecki", "wrzesi�ski", _
            "�remski", "jaroci�ski", "�redzki (wielkopolskie)", "leszczy�ski", "gosty�ski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszy�ski", "ostrzeszowski", "k�pi�ski", "turecki", "ko�cia�ski")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent5
        End With
        
        With ActiveSheet.Shapes.Range(Array("Pozna�", "Konin", "Kalisz", "Leszno", "z�otowski", "pilski", "nakielski", _
            "chodzieski", "w�growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "mi�dzychodzki", "nowotomyski", "wolszty�ski", "grodziski (wielkopolskie)", _
            "pozna�ski", "gnie�nie�ski", "kolski", "koni�ski", "s�upecki", "wrzesi�ski", _
            "�remski", "jaroci�ski", "�redzki (wielkopolskie)", "leszczy�ski", "gosty�ski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszy�ski", "ostrzeszowski", "k�pi�ski", "turecki", "ko�cia�ski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent5
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "Gorz�w Wielkopolski", "Zielona G�ra", "gorzowski", _
            "sul�ci�ski", "s�ubicki", "mi�dzyrzecki", "�wiebodzi�ski", "zielonog�rski", _
            "kro�nie�ski (lubuskie)", "�arski", "nowosolski", "�aga�ski", "wschowski")). _
            Group.Name = "lubuskie"
        
        With ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "Gorz�w Wielkopolski", "Zielona G�ra", "gorzowski", _
            "sul�ci�ski", "s�ubicki", "mi�dzyrzecki", "�wiebodzi�ski", "zielonog�rski", _
            "kro�nie�ski (lubuskie)", "�arski", "nowosolski", "�aga�ski", "wschowski")).Fill
            
            .ForeColor.RGB = RGB(165, 27, 96)
        End With
        
        With ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "Gorz�w Wielkopolski", "Zielona G�ra", "gorzowski", _
            "sul�ci�ski", "s�ubicki", "mi�dzyrzecki", "�wiebodzi�ski", "zielonog�rski", _
            "kro�nie�ski (lubuskie)", "�arski", "nowosolski", "�aga�ski", "wschowski")).Line
            
            .ForeColor.RGB = RGB(165, 27, 96)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Wroc�aw", "Jelenia G�ra", "Legnica", "wa�brzyski", "kamiennog�rski", "z�otoryjski" _
            , "luba�ski", "zgorzelecki", "boles�awiecki", "polkowicki", "g�ogowski", _
            "g�rowski", "wo�owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "ole�nicki", "wroc�awski", "o�awski", "strzeli�ski", "dzier�oniowski", _
            "z�bkowicki", "lubi�ski", "�redzki (dolny�l�sk)", "�widnicki (dolno�l�skie)", _
            "lw�wecki", "jeleniog�rski", "k�odzki")).Group.Name = "dolno�l�skie"
    
        With ActiveSheet.Shapes.Range(Array("Wroc�aw", "Jelenia G�ra", "Legnica", "wa�brzyski", "kamiennog�rski", "z�otoryjski" _
            , "luba�ski", "zgorzelecki", "boles�awiecki", "polkowicki", "g�ogowski", _
            "g�rowski", "wo�owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "ole�nicki", "wroc�awski", "o�awski", "strzeli�ski", "dzier�oniowski", _
            "z�bkowicki", "lubi�ski", "�redzki (dolny�l�sk)", "�widnicki (dolno�l�skie)", _
            "lw�wecki", "jeleniog�rski", "k�odzki")).Fill
            
            .ForeColor.RGB = RGB(255, 0, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Wroc�aw", "Jelenia G�ra", "Legnica", "wa�brzyski", "kamiennog�rski", "z�otoryjski" _
            , "luba�ski", "zgorzelecki", "boles�awiecki", "polkowicki", "g�ogowski", _
            "g�rowski", "wo�owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "ole�nicki", "wroc�awski", "o�awski", "strzeli�ski", "dzier�oniowski", _
            "z�bkowicki", "lubi�ski", "�redzki (dolny�l�sk)", "�widnicki (dolno�l�skie)", _
            "lw�wecki", "jeleniog�rski", "k�odzki")).Line
            
            .ForeColor.RGB = RGB(255, 0, 0)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Suwa�ki", "Bia�ystok", "�om�a", "suwalski", "sejne�ski", "augustowski", _
            "grajewski", "kolne�ski", "moniecki", "sok�lski", "bia�ostocki", "�om�y�ski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Group.Name = "podlaskie"
        
        With ActiveSheet.Shapes.Range(Array("Suwa�ki", "Bia�ystok", "�om�a", "suwalski", "sejne�ski", "augustowski", _
            "grajewski", "kolne�ski", "moniecki", "sok�lski", "bia�ostocki", "�om�y�ski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        End With
        
        With ActiveSheet.Shapes.Range(Array("Suwa�ki", "Bia�ystok", "�om�a", "suwalski", "sejne�ski", "augustowski", _
            "grajewski", "kolne�ski", "moniecki", "sok�lski", "bia�ostocki", "�om�y�ski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Ostro��ka", "P�ock", "Warszawa", "Radom", "Siedlce", "ostro��cki", "przasnyski", "ciechanowski", _
            "m�awski", "�uromi�ski", "sierpecki", "gostyni�ski", "p�ocki", "p�o�ski", _
            "pu�tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko�owski", "�osicki", "siedlecki", "w�growski", _
            "wo�omi�ski", "mi�ski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "�yrardowski", "piaseczy�ski", "otwocki", "gr�jecki", "przysuski", _
            "bia�obrzeski", "kozienicki", "garwoli�ski", "zwole�ski", "szyd�owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Group.Name = "mazowieckie"
            
        With ActiveSheet.Shapes.Range(Array("Ostro��ka", "P�ock", "Warszawa", "Radom", "Siedlce", "ostro��cki", "przasnyski", "ciechanowski", _
            "m�awski", "�uromi�ski", "sierpecki", "gostyni�ski", "p�ocki", "p�o�ski", _
            "pu�tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko�owski", "�osicki", "siedlecki", "w�growski", _
            "wo�omi�ski", "mi�ski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "�yrardowski", "piaseczy�ski", "otwocki", "gr�jecki", "przysuski", _
            "bia�obrzeski", "kozienicki", "garwoli�ski", "zwole�ski", "szyd�owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Fill
            
            .ForeColor.RGB = RGB(255, 204, 51)
            
         End With
         
        With ActiveSheet.Shapes.Range(Array("Ostro��ka", "P�ock", "Warszawa", "Radom", "Siedlce", "ostro��cki", "przasnyski", "ciechanowski", _
            "m�awski", "�uromi�ski", "sierpecki", "gostyni�ski", "p�ocki", "p�o�ski", _
            "pu�tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko�owski", "�osicki", "siedlecki", "w�growski", _
            "wo�omi�ski", "mi�ski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "�yrardowski", "piaseczy�ski", "otwocki", "gr�jecki", "przysuski", _
            "bia�obrzeski", "kozienicki", "garwoli�ski", "zwole�ski", "szyd�owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Line
            
            .ForeColor.RGB = RGB(255, 204, 51)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("rawski", "��d�", "Piotrk�w Trybunalski", "Skierniewice", "skierniewicki", "�owicki", _
            "kutnowski", "��czycki", "zgierski", "brzezi�ski", "tomaszowski (��dzkie)", _
            "opoczy�ski", "��dzki wschodni", "piotrkowski", "radomszcza�ski", "paj�cza�ski" _
            , "be�chatowski", "pabianicki", "�aski", "zdu�skowolski", "podd�bicki", _
            "sieradzki", "wielu�ski", "wieruszowski")).Group.Name = "��dzkie"
        
        With ActiveSheet.Shapes.Range(Array("rawski", "��d�", "Piotrk�w Trybunalski", "Skierniewice", "skierniewicki", "�owicki", _
            "kutnowski", "��czycki", "zgierski", "brzezi�ski", "tomaszowski (��dzkie)", _
            "opoczy�ski", "��dzki wschodni", "piotrkowski", "radomszcza�ski", "paj�cza�ski" _
            , "be�chatowski", "pabianicki", "�aski", "zdu�skowolski", "podd�bicki", _
            "sieradzki", "wielu�ski", "wieruszowski")).Fill
            
            .ForeColor.RGB = RGB(0, 176, 80)
        End With
        
        With ActiveSheet.Shapes.Range(Array("rawski", "��d�", "Piotrk�w Trybunalski", "Skierniewice", "skierniewicki", "�owicki", _
            "kutnowski", "��czycki", "zgierski", "brzezi�ski", "tomaszowski (��dzkie)", _
            "opoczy�ski", "��dzki wschodni", "piotrkowski", "radomszcza�ski", "paj�cza�ski" _
            , "be�chatowski", "pabianicki", "�aski", "zdu�skowolski", "podd�bicki", _
            "sieradzki", "wielu�ski", "wieruszowski")).Line
            
            .ForeColor.RGB = RGB(0, 176, 80)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Opole", "g�ubczycki", "k�dzierzy�sko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys�owski", "brzeski (opolskie)", "nyski")).Group.Name = "opolskie"
        
        With ActiveSheet.Shapes.Range(Array("Opole", "g�ubczycki", "k�dzierzy�sko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys�owski", "brzeski (opolskie)", "nyski")).Fill
            
            .ForeColor.RGB = RGB(0, 0, 204)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Opole", "g�ubczycki", "k�dzierzy�sko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys�owski", "brzeski (opolskie)", "nyski")).Line
            
            .ForeColor.RGB = RGB(0, 0, 204)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Cz�stochowa", "Bielsko-Bia�a", "Ruda �l�ska", "�wi�toch�owice", "Chorz�w" _
        , "Siemianowice �l�skie", "Piekary �l�skie", "Bytom", "Zabrze", "Gliwice", "Rybnik", "Jastrz�bie-Zdr�j", "�ory", _
        "Tychy", "Jaworzno", "Mys�owice", "Katowice", "Sosnowiec", "D�browa-G�rnicza", "bieru�sko-l�dzi�ski", "pszczy�ski", _
        "bielski (�l�skie)", "cieszy�ski", "wodzis�awski", "rybnicki", "raciborski", _
        "gliwicki", "miko�owski", "b�dzi�ski", "tarnog�rski", "zawiercia�ski", _
        "myszkowski", "cz�stochowski", "lubliniecki", "k�obucki", "�ywiecki")).Group.Name = "�l�skie"
    
        With ActiveSheet.Shapes.Range(Array("Cz�stochowa", "Bielsko-Bia�a", "Ruda �l�ska", "�wi�toch�owice", "Chorz�w" _
        , "Siemianowice �l�skie", "Piekary �l�skie", "Bytom", "Zabrze", "Gliwice", "Rybnik", "Jastrz�bie-Zdr�j", "�ory", _
        "Tychy", "Jaworzno", "Mys�owice", "Katowice", "Sosnowiec", "D�browa-G�rnicza", "bieru�sko-l�dzi�ski", "pszczy�ski", _
        "bielski (�l�skie)", "cieszy�ski", "wodzis�awski", "rybnicki", "raciborski", _
        "gliwicki", "miko�owski", "b�dzi�ski", "tarnog�rski", "zawiercia�ski", _
        "myszkowski", "cz�stochowski", "lubliniecki", "k�obucki", "�ywiecki")).Fill
        
            .ForeColor.RGB = RGB(204, 153, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Cz�stochowa", "Bielsko-Bia�a", "Ruda �l�ska", "�wi�toch�owice", "Chorz�w" _
        , "Siemianowice �l�skie", "Piekary �l�skie", "Bytom", "Zabrze", "Gliwice", "Rybnik", "Jastrz�bie-Zdr�j", "�ory", _
        "Tychy", "Jaworzno", "Mys�owice", "Katowice", "Sosnowiec", "D�browa-G�rnicza", "bieru�sko-l�dzi�ski", "pszczy�ski", _
        "bielski (�l�skie)", "cieszy�ski", "wodzis�awski", "rybnicki", "raciborski", _
        "gliwicki", "miko�owski", "b�dzi�ski", "tarnog�rski", "zawiercia�ski", _
        "myszkowski", "cz�stochowski", "lubliniecki", "k�obucki", "�ywiecki")).Line
        
             .ForeColor.RGB = RGB(204, 153, 0)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Kielce", "konecki", "skar�yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w�oszczowski", "j�drzejowski", "pi�czowski", "buski", "kazimierski")).Group.Name = "�wi�tokrzyskie"
            
        With ActiveSheet.Shapes.Range(Array("Kielce", "konecki", "skar�yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w�oszczowski", "j�drzejowski", "pi�czowski", "buski", "kazimierski")).Fill
            
            .ForeColor.RGB = RGB(204, 51, 255)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Kielce", "konecki", "skar�yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w�oszczowski", "j�drzejowski", "pi�czowski", "buski", "kazimierski")).Line
            
            .ForeColor.RGB = RGB(204, 51, 255)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Zamo��", "Che�m", "Bia�a Podlaska", "Lublin", "bialski", "radzy�ski", "�ukowski", "rycki", _
            "parczewski", "w�odawski", "lubartowski", "pu�awski", "opolski (lubelskie)", _
            "lubelski", "kra�nicki", "�widnicki (lubelskie)", "che�mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi�gorajski", "��czy�ski", _
            "tomaszowski (lubelskie)")).Group.Name = "lubelskie"
            
        With ActiveSheet.Shapes.Range(Array("Zamo��", "Che�m", "Bia�a Podlaska", "Lublin", "bialski", "radzy�ski", "�ukowski", "rycki", _
            "parczewski", "w�odawski", "lubartowski", "pu�awski", "opolski (lubelskie)", _
            "lubelski", "kra�nicki", "�widnicki (lubelskie)", "che�mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi�gorajski", "��czy�ski", _
            "tomaszowski (lubelskie)")).Fill
            
            .ForeColor.RGB = RGB(102, 0, 204)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Zamo��", "Che�m", "Bia�a Podlaska", "Lublin", "bialski", "radzy�ski", "�ukowski", "rycki", _
            "parczewski", "w�odawski", "lubartowski", "pu�awski", "opolski (lubelskie)", _
            "lubelski", "kra�nicki", "�widnicki (lubelskie)", "che�mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi�gorajski", "��czy�ski", _
            "tomaszowski (lubelskie)")).Line
            
            .ForeColor.RGB = RGB(102, 0, 204)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Rzesz�w", "Przemy�l", "Krosno", "Tarnobrzeg", "jaros�awski", "przemyski", "le�ajski", _
            "przeworski", "�a�cucki", "ni�a�ski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-s�dziszowski", "mielecki", "d�bicki", _
            "brzozowski", "kro�nie�ski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy�owski", "lubaczowski")).Group.Name = "podkarpackie"
        
        With ActiveSheet.Shapes.Range(Array("Rzesz�w", "Przemy�l", "Krosno", "Tarnobrzeg", "jaros�awski", "przemyski", "le�ajski", _
            "przeworski", "�a�cucki", "ni�a�ski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-s�dziszowski", "mielecki", "d�bicki", _
            "brzozowski", "kro�nie�ski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy�owski", "lubaczowski")).Fill
            
            .ForeColor.RGB = RGB(255, 64, 64)
            
        End With
        
        With ActiveSheet.Shapes.Range(Array("Rzesz�w", "Przemy�l", "Krosno", "Tarnobrzeg", "jaros�awski", "przemyski", "le�ajski", _
            "przeworski", "�a�cucki", "ni�a�ski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-s�dziszowski", "mielecki", "d�bicki", _
            "brzozowski", "kro�nie�ski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy�owski", "lubaczowski")).Line
            
            .ForeColor.RGB = RGB(255, 64, 64)
            .Weight = 2.5
        End With

        ActiveSheet.Shapes.Range(Array("Krak�w", "Tarn�w", "Nowy S�cz", "d�browski", "gorlicki", "nowos�decki", _
            "boche�ski", "brzeski (ma�opolskie)", "tarnowski", "limanowski", "wielicki", _
            "my�lenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "o�wi�cimski", "chrzanowski", "okulski", "nowotarski", "tatrza�ski")).Group.Name = "ma�opolskie"
                
        With ActiveSheet.Shapes.Range(Array("Krak�w", "Tarn�w", "Nowy S�cz", "d�browski", "gorlicki", "nowos�decki", _
            "boche�ski", "brzeski (ma�opolskie)", "tarnowski", "limanowski", "wielicki", _
            "my�lenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "o�wi�cimski", "chrzanowski", "okulski", "nowotarski", "tatrza�ski")).Fill
            
            .ForeColor.RGB = RGB(204, 51, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Krak�w", "Tarn�w", "Nowy S�cz", "d�browski", "gorlicki", "nowos�decki", _
            "boche�ski", "brzeski (ma�opolskie)", "tarnowski", "limanowski", "wielicki", _
            "my�lenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "o�wi�cimski", "chrzanowski", "okulski", "nowotarski", "tatrza�ski")).Line
            
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
    
            .ChartTitle.Text = "Wojew�dztwo mazowieckie - prognoza liczby ludno�ci na lata 2011-2035"
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
        
        ActiveSheet.Shapes.Range(Array("Ruda �l�ska ", "�wi�toch�owice ", "Chorz�w " _
        , "b�dzi�ski ", "Siemianowice �l�skie ", "Piekary �l�skie ", "Bytom ", _
        "Zabrze ", "Gliwice ", "Rybnik ", "Jastrz�bie-Zdr�j ", "miko�owski ", "�ory ", _
        "Tychy ", "bieru�sko-l�dzi�ski ", "Jaworzno ", "Mys�owice ", "Katowice ", _
        "Sosnowiec ", "D�browa-G�rnicza ")).Group.Name = " "
        
         On Error Resume Next
        
        ActiveSheet.Shapes.Range(Array(" ")).Visible = msoFalse
    
        With ActiveSheet.Shapes.Range(Array("Woj")).Glow
        
            .color.ObjectThemeColor = msoThemeColorBackground1
            ' .Transparency = 0.2
            .Radius = 5
        End With
              
        ActiveSheet.Shapes.Range(Array("pucki", "Gdynia", "Sopot", "Gda�sk", "S�upsk", "wejherowski", "l�borski", "s�upski" _
            , "bytowski", "cz�uchowski", "chojnicki", "ko�cierski", "kartuski", "gda�ski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzy�ski")).Group.Name = "pomorskie"
            
        With ActiveSheet.Shapes.Range(Array("pucki", "Gdynia", "Sopot", "Gda�sk", "S�upsk", "wejherowski", "l�borski", "s�upski" _
            , "bytowski", "cz�uchowski", "chojnicki", "ko�cierski", "kartuski", "gda�ski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzy�ski")).Fill
            
            .ForeColor.RGB = RGB(255, 192, 0)
            .Solid
        End With
        
        With ActiveSheet.Shapes.Range(Array("pucki", "Gdynia", "Sopot", "Gda�sk", "S�upsk", "wejherowski", "l�borski", "s�upski" _
            , "bytowski", "cz�uchowski", "chojnicki", "ko�cierski", "kartuski", "gda�ski", _
            "starogardzki", "tczewski", "nowodworski (pomorskie)", "malborski", "sztumski" _
            , "kwidzy�ski")).Line
            
            .ForeColor.RGB = RGB(255, 192, 0)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("s�awie�ski", "Szczecin", "Koszalin", "bia�ogardzki", "koszali�ski", _
            "szczecinecki", "wa�ecki", "drawski", "�widwi�ski", "ko�obrzeski", "gryficki", _
            "�obeski", "�winouj�cie", "goleniowski", "stargardzki", "choszcze�ski", _
            "pyrzycki", "my�liborski", "gryfi�ski", "policki", "kamie�ski")).Group.Name = "zachodniopomorskie"
    
        With ActiveSheet.Shapes.Range(Array("s�awie�ski", "Szczecin", "Koszalin", "bia�ogardzki", "koszali�ski", _
            "szczecinecki", "wa�ecki", "drawski", "�widwi�ski", "ko�obrzeski", "gryficki", _
            "�obeski", "�winouj�cie", "goleniowski", "stargardzki", "choszcze�ski", _
            "pyrzycki", "my�liborski", "gryfi�ski", "policki", "kamie�ski")).Fill
            .ForeColor.RGB = RGB(192, 0, 0)
        End With

        With ActiveSheet.Shapes.Range(Array("s�awie�ski", "Szczecin", "Koszalin", "bia�ogardzki", "koszali�ski", _
            "szczecinecki", "wa�ecki", "drawski", "�widwi�ski", "ko�obrzeski", "gryficki", _
            "�obeski", "�winouj�cie", "goleniowski", "stargardzki", "choszcze�ski", _
            "pyrzycki", "my�liborski", "gryfi�ski", "policki", "kamie�ski")).Line
            .ForeColor.RGB = RGB(192, 0, 0)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Elbl�g", "Olsztyn", "braniewski", "elbl�ski", "i�awski", _
            "nowomiejski", "dzia�dowski", "ostr�dzki", "lidzbarski", "bartoszycki", _
            "olszty�ski", "nidzicki", "szczycie�ski", "mr�gowski", "k�trzy�ski", _
            "w�gorzewski", "go�dapski", "olecki", "gi�ycki", "e�cki", "piski")).Group.Name = "warmi�sko-mazurskie"
    
        With ActiveSheet.Shapes.Range(Array("Elbl�g", "Olsztyn", "braniewski", "elbl�ski", "i�awski", _
            "nowomiejski", "dzia�dowski", "ostr�dzki", "lidzbarski", "bartoszycki", _
            "olszty�ski", "nidzicki", "szczycie�ski", "mr�gowski", "k�trzy�ski", _
            "w�gorzewski", "go�dapski", "olecki", "gi�ycki", "e�cki", "piski")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent3
        End With
        
        With ActiveSheet.Shapes.Range(Array("Elbl�g", "Olsztyn", "braniewski", "elbl�ski", "i�awski", _
            "nowomiejski", "dzia�dowski", "ostr�dzki", "lidzbarski", "bartoszycki", _
            "olszty�ski", "nidzicki", "szczycie�ski", "mr�gowski", "k�trzy�ski", _
            "w�gorzewski", "go�dapski", "olecki", "gi�ycki", "e�cki", "piski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent3
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("s�pole�ski", "tucholski", "Bydgoszcz", "Toru�", "Grudzi�dz", "W�oc�awek", "che�mi�ski", _
            "�wiecki", "grudzi�dzki", "w�brzeski", "brodnicki", "rypi�ski", _
            "golubsko-dobrzy�ski", "lipnowski", "w�oc�awski", "radziejowski", _
            "aleksandrowski", "toru�ski", "inowroc�awski", "mogile�ski", "bydgoski", _
            "�ni�ski")).Group.Name = "kujawsko-pomorskie"
            
        With ActiveSheet.Shapes.Range(Array("s�pole�ski", "tucholski", "Bydgoszcz", "Toru�", "Grudzi�dz", "W�oc�awek", "che�mi�ski", _
            "�wiecki", "grudzi�dzki", "w�brzeski", "brodnicki", "rypi�ski", _
            "golubsko-dobrzy�ski", "lipnowski", "w�oc�awski", "radziejowski", _
            "aleksandrowski", "toru�ski", "inowroc�awski", "mogile�ski", "bydgoski", _
            "�ni�ski")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent4
        End With
        
        With ActiveSheet.Shapes.Range(Array("s�pole�ski", "tucholski", "Bydgoszcz", "Toru�", "Grudzi�dz", "W�oc�awek", "che�mi�ski", _
            "�wiecki", "grudzi�dzki", "w�brzeski", "brodnicki", "rypi�ski", _
            "golubsko-dobrzy�ski", "lipnowski", "w�oc�awski", "radziejowski", _
            "aleksandrowski", "toru�ski", "inowroc�awski", "mogile�ski", "bydgoski", _
            "�ni�ski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent4
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Pozna�", "Konin", "Kalisz", "Leszno", "z�otowski", "pilski", "nakielski", _
            "chodzieski", "w�growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "mi�dzychodzki", "nowotomyski", "wolszty�ski", "grodziski (wielkopolskie)", _
            "pozna�ski", "gnie�nie�ski", "kolski", "koni�ski", "s�upecki", "wrzesi�ski", _
            "�remski", "jaroci�ski", "�redzki (wielkopolskie)", "leszczy�ski", "gosty�ski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszy�ski", "ostrzeszowski", "k�pi�ski", "turecki", "ko�cia�ski")).Group.Name = "wielkopolskie"
    
        With ActiveSheet.Shapes.Range(Array("Pozna�", "Konin", "Kalisz", "Leszno", "z�otowski", "pilski", "nakielski", _
            "chodzieski", "w�growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "mi�dzychodzki", "nowotomyski", "wolszty�ski", "grodziski (wielkopolskie)", _
            "pozna�ski", "gnie�nie�ski", "kolski", "koni�ski", "s�upecki", "wrzesi�ski", _
            "�remski", "jaroci�ski", "�redzki (wielkopolskie)", "leszczy�ski", "gosty�ski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszy�ski", "ostrzeszowski", "k�pi�ski", "turecki", "ko�cia�ski")).Fill
            .ForeColor.ObjectThemeColor = msoThemeColorAccent5
        End With
        
        With ActiveSheet.Shapes.Range(Array("Pozna�", "Konin", "Kalisz", "Leszno", "z�otowski", "pilski", "nakielski", _
            "chodzieski", "w�growiecki", "obornicki", "czarnkowsko-trzcianecki", _
            "szamotulski", "mi�dzychodzki", "nowotomyski", "wolszty�ski", "grodziski (wielkopolskie)", _
            "pozna�ski", "gnie�nie�ski", "kolski", "koni�ski", "s�upecki", "wrzesi�ski", _
            "�remski", "jaroci�ski", "�redzki (wielkopolskie)", "leszczy�ski", "gosty�ski" _
            , "rawicki", "pleszewski", "kaliski", "ostrowski (wielkopolskie)", _
            "krotoszy�ski", "ostrzeszowski", "k�pi�ski", "turecki", "ko�cia�ski")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent5
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "Gorz�w Wielkopolski", "Zielona G�ra", "gorzowski", _
            "sul�ci�ski", "s�ubicki", "mi�dzyrzecki", "�wiebodzi�ski", "zielonog�rski", _
            "kro�nie�ski (lubuskie)", "�arski", "nowosolski", "�aga�ski", "wschowski")). _
            Group.Name = "lubuskie"
        
        With ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "Gorz�w Wielkopolski", "Zielona G�ra", "gorzowski", _
            "sul�ci�ski", "s�ubicki", "mi�dzyrzecki", "�wiebodzi�ski", "zielonog�rski", _
            "kro�nie�ski (lubuskie)", "�arski", "nowosolski", "�aga�ski", "wschowski")).Fill
            
            .ForeColor.RGB = RGB(165, 27, 96)
        End With
        
        With ActiveSheet.Shapes.Range(Array("strzelecko-drezdenecki", "Gorz�w Wielkopolski", "Zielona G�ra", "gorzowski", _
            "sul�ci�ski", "s�ubicki", "mi�dzyrzecki", "�wiebodzi�ski", "zielonog�rski", _
            "kro�nie�ski (lubuskie)", "�arski", "nowosolski", "�aga�ski", "wschowski")).Line
            .ForeColor.RGB = RGB(165, 27, 96)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Wroc�aw", "Jelenia G�ra", "Legnica", "wa�brzyski", "kamiennog�rski", "z�otoryjski" _
            , "luba�ski", "zgorzelecki", "boles�awiecki", "polkowicki", "g�ogowski", _
            "g�rowski", "wo�owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "ole�nicki", "wroc�awski", "o�awski", "strzeli�ski", "dzier�oniowski", _
            "z�bkowicki", "lubi�ski", "�redzki (dolny�l�sk)", "�widnicki (dolno�l�skie)", _
            "lw�wecki", "jeleniog�rski", "k�odzki")).Group.Name = "dolno�l�skie"
    
        With ActiveSheet.Shapes.Range(Array("Wroc�aw", "Jelenia G�ra", "Legnica", "wa�brzyski", "kamiennog�rski", "z�otoryjski" _
            , "luba�ski", "zgorzelecki", "boles�awiecki", "polkowicki", "g�ogowski", _
            "g�rowski", "wo�owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "ole�nicki", "wroc�awski", "o�awski", "strzeli�ski", "dzier�oniowski", _
            "z�bkowicki", "lubi�ski", "�redzki (dolny�l�sk)", "�widnicki (dolno�l�skie)", _
            "lw�wecki", "jeleniog�rski", "k�odzki")).Fill
            
            .ForeColor.RGB = RGB(255, 0, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Wroc�aw", "Jelenia G�ra", "Legnica", "wa�brzyski", "kamiennog�rski", "z�otoryjski" _
            , "luba�ski", "zgorzelecki", "boles�awiecki", "polkowicki", "g�ogowski", _
            "g�rowski", "wo�owski", "legnicki", "jaworski", "trzebnicki", "milicki", _
            "ole�nicki", "wroc�awski", "o�awski", "strzeli�ski", "dzier�oniowski", _
            "z�bkowicki", "lubi�ski", "�redzki (dolny�l�sk)", "�widnicki (dolno�l�skie)", _
            "lw�wecki", "jeleniog�rski", "k�odzki")).Line
            
            .ForeColor.RGB = RGB(255, 0, 0)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Suwa�ki", "Bia�ystok", "�om�a", "suwalski", "sejne�ski", "augustowski", _
            "grajewski", "kolne�ski", "moniecki", "sok�lski", "bia�ostocki", "�om�y�ski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Group.Name = "podlaskie"
        
        With ActiveSheet.Shapes.Range(Array("Suwa�ki", "Bia�ystok", "�om�a", "suwalski", "sejne�ski", "augustowski", _
            "grajewski", "kolne�ski", "moniecki", "sok�lski", "bia�ostocki", "�om�y�ski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Fill
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
        End With
        
        With ActiveSheet.Shapes.Range(Array("Suwa�ki", "Bia�ystok", "�om�a", "suwalski", "sejne�ski", "augustowski", _
            "grajewski", "kolne�ski", "moniecki", "sok�lski", "bia�ostocki", "�om�y�ski", _
            "zambrowski", "wysokomazowiecki", "bielski (podlaskie)", "hajnowski", _
            "siemiatycki")).Line
            
            .ForeColor.ObjectThemeColor = msoThemeColorAccent1
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Ostro��ka", "P�ock", "Warszawa", "Radom", "Siedlce", "ostro��cki", "przasnyski", "ciechanowski", _
            "m�awski", "�uromi�ski", "sierpecki", "gostyni�ski", "p�ocki", "p�o�ski", _
            "pu�tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko�owski", "�osicki", "siedlecki", "w�growski", _
            "wo�omi�ski", "mi�ski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "�yrardowski", "piaseczy�ski", "otwocki", "gr�jecki", "przysuski", _
            "bia�obrzeski", "kozienicki", "garwoli�ski", "zwole�ski", "szyd�owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Group.Name = "mazowieckie"
            
        With ActiveSheet.Shapes.Range(Array("Ostro��ka", "P�ock", "Warszawa", "Radom", "Siedlce", "ostro��cki", "przasnyski", "ciechanowski", _
            "m�awski", "�uromi�ski", "sierpecki", "gostyni�ski", "p�ocki", "p�o�ski", _
            "pu�tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko�owski", "�osicki", "siedlecki", "w�growski", _
            "wo�omi�ski", "mi�ski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "�yrardowski", "piaseczy�ski", "otwocki", "gr�jecki", "przysuski", _
            "bia�obrzeski", "kozienicki", "garwoli�ski", "zwole�ski", "szyd�owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Fill
            
            .ForeColor.RGB = RGB(255, 204, 51)
         End With
         
        With ActiveSheet.Shapes.Range(Array("Ostro��ka", "P�ock", "Warszawa", "Radom", "Siedlce", "ostro��cki", "przasnyski", "ciechanowski", _
            "m�awski", "�uromi�ski", "sierpecki", "gostyni�ski", "p�ocki", "p�o�ski", _
            "pu�tuski", "nowodworski (mazowieckie)", "sochaczewski", "wyszkowski", _
            "ostrowski (mazowieckie)", "soko�owski", "�osicki", "siedlecki", "w�growski", _
            "wo�omi�ski", "mi�ski", "warszawski zachodni", "pruszkowski", "grodziski (mazowieckie)", _
            "�yrardowski", "piaseczy�ski", "otwocki", "gr�jecki", "przysuski", _
            "bia�obrzeski", "kozienicki", "garwoli�ski", "zwole�ski", "szyd�owiecki", _
            "makowski", "legionowski", "radomski", "lipski")).Line
            
            .ForeColor.RGB = RGB(255, 204, 51)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("rawski", "��d�", "Piotrk�w Trybunalski", "Skierniewice", "skierniewicki", "�owicki", _
            "kutnowski", "��czycki", "zgierski", "brzezi�ski", "tomaszowski (��dzkie)", _
            "opoczy�ski", "��dzki wschodni", "piotrkowski", "radomszcza�ski", "paj�cza�ski" _
            , "be�chatowski", "pabianicki", "�aski", "zdu�skowolski", "podd�bicki", _
            "sieradzki", "wielu�ski", "wieruszowski")).Group.Name = "��dzkie"
        
        With ActiveSheet.Shapes.Range(Array("rawski", "��d�", "Piotrk�w Trybunalski", "Skierniewice", "skierniewicki", "�owicki", _
            "kutnowski", "��czycki", "zgierski", "brzezi�ski", "tomaszowski (��dzkie)", _
            "opoczy�ski", "��dzki wschodni", "piotrkowski", "radomszcza�ski", "paj�cza�ski" _
            , "be�chatowski", "pabianicki", "�aski", "zdu�skowolski", "podd�bicki", _
            "sieradzki", "wielu�ski", "wieruszowski")).Fill
            .ForeColor.RGB = RGB(0, 176, 80)
        End With
        
        With ActiveSheet.Shapes.Range(Array("rawski", "��d�", "Piotrk�w Trybunalski", "Skierniewice", "skierniewicki", "�owicki", _
            "kutnowski", "��czycki", "zgierski", "brzezi�ski", "tomaszowski (��dzkie)", _
            "opoczy�ski", "��dzki wschodni", "piotrkowski", "radomszcza�ski", "paj�cza�ski" _
            , "be�chatowski", "pabianicki", "�aski", "zdu�skowolski", "podd�bicki", _
            "sieradzki", "wielu�ski", "wieruszowski")).Line
            
            .ForeColor.RGB = RGB(0, 176, 80)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Opole", "g�ubczycki", "k�dzierzy�sko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys�owski", "brzeski (opolskie)", "nyski")).Group.Name = "opolskie"
        
        With ActiveSheet.Shapes.Range(Array("Opole", "g�ubczycki", "k�dzierzy�sko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys�owski", "brzeski (opolskie)", "nyski")).Fill
            
            .ForeColor.RGB = RGB(0, 0, 204)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Opole", "g�ubczycki", "k�dzierzy�sko-kozielski", _
            "strzelecki", "oleski", "kluczborski", "krapkowicki", "prudnicki", _
            "opolski (opolskie)", "namys�owski", "brzeski (opolskie)", "nyski")).Line
            
            .ForeColor.RGB = RGB(0, 0, 204)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Cz�stochowa", "Bielsko-Bia�a", "Ruda �l�ska", "�wi�toch�owice", "Chorz�w" _
        , "Siemianowice �l�skie", "Piekary �l�skie", "Bytom", "Zabrze", "Gliwice", "Rybnik", "Jastrz�bie-Zdr�j", "�ory", _
        "Tychy", "Jaworzno", "Mys�owice", "Katowice", "Sosnowiec", "D�browa-G�rnicza", "bieru�sko-l�dzi�ski", "pszczy�ski", _
        "bielski (�l�skie)", "cieszy�ski", "wodzis�awski", "rybnicki", "raciborski", _
        "gliwicki", "miko�owski", "b�dzi�ski", "tarnog�rski", "zawiercia�ski", _
        "myszkowski", "cz�stochowski", "lubliniecki", "k�obucki", "�ywiecki")).Group.Name = "�l�skie"
    
        With ActiveSheet.Shapes.Range(Array("Cz�stochowa", "Bielsko-Bia�a", "Ruda �l�ska", "�wi�toch�owice", "Chorz�w" _
        , "Siemianowice �l�skie", "Piekary �l�skie", "Bytom", "Zabrze", "Gliwice", "Rybnik", "Jastrz�bie-Zdr�j", "�ory", _
        "Tychy", "Jaworzno", "Mys�owice", "Katowice", "Sosnowiec", "D�browa-G�rnicza", "bieru�sko-l�dzi�ski", "pszczy�ski", _
        "bielski (�l�skie)", "cieszy�ski", "wodzis�awski", "rybnicki", "raciborski", _
        "gliwicki", "miko�owski", "b�dzi�ski", "tarnog�rski", "zawiercia�ski", _
        "myszkowski", "cz�stochowski", "lubliniecki", "k�obucki", "�ywiecki")).Fill
            .ForeColor.RGB = RGB(204, 153, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Cz�stochowa", "Bielsko-Bia�a", "Ruda �l�ska", "�wi�toch�owice", "Chorz�w" _
        , "Siemianowice �l�skie", "Piekary �l�skie", "Bytom", "Zabrze", "Gliwice", "Rybnik", "Jastrz�bie-Zdr�j", "�ory", _
        "Tychy", "Jaworzno", "Mys�owice", "Katowice", "Sosnowiec", "D�browa-G�rnicza", "bieru�sko-l�dzi�ski", "pszczy�ski", _
        "bielski (�l�skie)", "cieszy�ski", "wodzis�awski", "rybnicki", "raciborski", _
        "gliwicki", "miko�owski", "b�dzi�ski", "tarnog�rski", "zawiercia�ski", _
        "myszkowski", "cz�stochowski", "lubliniecki", "k�obucki", "�ywiecki")).Line
             .ForeColor.RGB = RGB(204, 153, 0)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Kielce", "konecki", "skar�yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w�oszczowski", "j�drzejowski", "pi�czowski", "buski", "kazimierski")).Group.Name = "�wi�tokrzyskie"
            
        With ActiveSheet.Shapes.Range(Array("Kielce", "konecki", "skar�yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w�oszczowski", "j�drzejowski", "pi�czowski", "buski", "kazimierski")).Fill
            
            .ForeColor.RGB = RGB(204, 51, 255)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Kielce", "konecki", "skar�yski", "starachowicki", _
            "ostrowiecki", "opatowski", "sandomierski", "staszowski", "kielecki", _
            "w�oszczowski", "j�drzejowski", "pi�czowski", "buski", "kazimierski")).Line
            
            .ForeColor.RGB = RGB(204, 51, 255)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Zamo��", "Che�m", "Bia�a Podlaska", "Lublin", "bialski", "radzy�ski", "�ukowski", "rycki", _
            "parczewski", "w�odawski", "lubartowski", "pu�awski", "opolski (lubelskie)", _
            "lubelski", "kra�nicki", "�widnicki (lubelskie)", "che�mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi�gorajski", "��czy�ski", _
            "tomaszowski (lubelskie)")).Group.Name = "lubelskie"
            
        With ActiveSheet.Shapes.Range(Array("Zamo��", "Che�m", "Bia�a Podlaska", "Lublin", "bialski", "radzy�ski", "�ukowski", "rycki", _
            "parczewski", "w�odawski", "lubartowski", "pu�awski", "opolski (lubelskie)", _
            "lubelski", "kra�nicki", "�widnicki (lubelskie)", "che�mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi�gorajski", "��czy�ski", _
            "tomaszowski (lubelskie)")).Fill
            .ForeColor.RGB = RGB(102, 0, 204)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Zamo��", "Che�m", "Bia�a Podlaska", "Lublin", "bialski", "radzy�ski", "�ukowski", "rycki", _
            "parczewski", "w�odawski", "lubartowski", "pu�awski", "opolski (lubelskie)", _
            "lubelski", "kra�nicki", "�widnicki (lubelskie)", "che�mski", "krasnostawski", _
            "hrubieszowski", "zamojski", "janowski", "bi�gorajski", "��czy�ski", _
            "tomaszowski (lubelskie)")).Line
            .ForeColor.RGB = RGB(102, 0, 204)
            .Weight = 2.5
        End With
    
        ActiveSheet.Shapes.Range(Array("Rzesz�w", "Przemy�l", "Krosno", "Tarnobrzeg", "jaros�awski", "przemyski", "le�ajski", _
            "przeworski", "�a�cucki", "ni�a�ski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-s�dziszowski", "mielecki", "d�bicki", _
            "brzozowski", "kro�nie�ski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy�owski", "lubaczowski")).Group.Name = "podkarpackie"
        
        With ActiveSheet.Shapes.Range(Array("Rzesz�w", "Przemy�l", "Krosno", "Tarnobrzeg", "jaros�awski", "przemyski", "le�ajski", _
            "przeworski", "�a�cucki", "ni�a�ski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-s�dziszowski", "mielecki", "d�bicki", _
            "brzozowski", "kro�nie�ski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy�owski", "lubaczowski")).Fill
            
            .ForeColor.RGB = RGB(255, 64, 64)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Rzesz�w", "Przemy�l", "Krosno", "Tarnobrzeg", "jaros�awski", "przemyski", "le�ajski", _
            "przeworski", "�a�cucki", "ni�a�ski", "stalowowolski", "tarnobrzeski", _
            "kolbuszowski", "rzeszowski", "ropczycko-s�dziszowski", "mielecki", "d�bicki", _
            "brzozowski", "kro�nie�ski", "jasielski", "sanocki", "leski", "bieszczadzki", _
            "strzy�owski", "lubaczowski")).Line
            
            .ForeColor.RGB = RGB(255, 64, 64)
            .Weight = 2.5
        End With

        ActiveSheet.Shapes.Range(Array("Krak�w", "Tarn�w", "Nowy S�cz", "d�browski", "gorlicki", "nowos�decki", _
            "boche�ski", "brzeski (ma�opolskie)", "tarnowski", "limanowski", "wielicki", _
            "my�lenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "o�wi�cimski", "chrzanowski", "okulski", "nowotarski", "tatrza�ski")).Group.Name = "ma�opolskie"
                
        With ActiveSheet.Shapes.Range(Array("Krak�w", "Tarn�w", "Nowy S�cz", "d�browski", "gorlicki", "nowos�decki", _
            "boche�ski", "brzeski (ma�opolskie)", "tarnowski", "limanowski", "wielicki", _
            "my�lenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "o�wi�cimski", "chrzanowski", "okulski", "nowotarski", "tatrza�ski")).Fill
            
            .ForeColor.RGB = RGB(204, 51, 0)
        End With
        
        With ActiveSheet.Shapes.Range(Array("Krak�w", "Tarn�w", "Nowy S�cz", "d�browski", "gorlicki", "nowos�decki", _
            "boche�ski", "brzeski (ma�opolskie)", "tarnowski", "limanowski", "wielicki", _
            "my�lenicki", "suski", "proszowicki", "miechowski", "krakowski", "wadowicki", _
            "o�wi�cimski", "chrzanowski", "okulski", "nowotarski", "tatrza�ski")).Line
            
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
    
            .ChartTitle.Text = "Wojew�dztwo mazowieckie - prognoza liczby ludno�ci na lata 2011-2035"
            .SeriesCollection(1).Values = Worksheets("Mapka dane").Range("C" & 426 & "," & "F" & 426 & "," & "I" & 426 & "," & "L" & 426 & "," & "O" & 426 & "," & "R" & 426 & "," & "U" & 426 & "," & "X" & 426 & "," & "AA" & 426 & "," & "AD" & 426 & "," & "AG" & 426 & "," & "AJ" & 426 & "," & "AM" & 426 & "," & "AP" & 426 & "," & "AS" & 426 & "," & "AV" & 426 & "," & "AY" & 426 & "," & "BB" & 426 & "," & "BE" & 426 & "," & "BK" & 426 & "," & "BN" & 426 & "," & "BQ" & 426 & "," & "BT" & 426 & "," & "BW" & 426)
            .SeriesCollection(2).Values = Worksheets("Mapka dane").Range("D" & 426 & "," & "G" & 426 & "," & "J" & 426 & "," & "M" & 426 & "," & "P" & 426 & "," & "S" & 426 & "," & "V" & 426 & "," & "Y" & 426 & "," & "AB" & 426 & "," & "AE" & 426 & "," & "AH" & 426 & "," & "AK" & 426 & "," & "AN" & 426 & "," & "AQ" & 426 & "," & "AT" & 426 & "," & "AW" & 426 & "," & "AZ" & 426 & "," & "BC" & 426 & "," & "BF" & 426 & "," & "BL" & 426 & "," & "BO" & 426 & "," & "BR" & 426 & "," & "BU" & 426 & "," & "BX" & 426)
            .SeriesCollection(3).Values = Worksheets("Mapka dane").Range("E" & 426 & "," & "H" & 426 & "," & "K" & 426 & "," & "N" & 426 & "," & "Q" & 426 & "," & "T" & 426 & "," & "W" & 426 & "," & "Z" & 426 & "," & "AC" & 426 & "," & "AF" & 426 & "," & "AI" & 426 & "," & "AL" & 426 & "," & "AO" & 426 & "," & "AR" & 426 & "," & "AU" & 426 & "," & "AX" & 426 & "," & "BA" & 426 & "," & "BD" & 426 & "," & "BG" & 426 & "," & "BM" & 426 & "," & "BP" & 426 & "," & "BS" & 426 & "," & "BV" & 426 & "," & "BY" & 426)
    
        End With
        
    End If
    
    RefreshCharts  'odswieza wykresy, �eby update'owa�y si� w czasie dzia�ania excela

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
