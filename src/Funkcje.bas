Attribute VB_Name = "Funkcje"
Option Explicit

Public counter As String
Public mapa(1 To 57, 1 To 70, 0 To 5) As String
Public numer(1 To 57, 1 To 70, 0 To 5) As Integer
Public Woj(1 To 57, 1 To 70, 0 To 5) As String
Public WojNum(1 To 57, 1 To 70, 0 To 5) As Integer


Public Function TworzenieObszaru()

    Dim i, j As Integer
    
    For i = 1 To 70
    
        For j = 1 To 57
    
            Sheets("Powiaty").Cells(j, i).FormulaR1C1 = "=IFERROR(HYPERLINK(Wspolrzedne(ROW(),COLUMN()),""""),"""")"
        
        Next j
        
    Next i

End Function

Public Function Wspolrzedne(row As Integer, column As Integer)

    Dim pozname, title As String
    Dim num As Integer
    
    If ActiveSheet.Shapes.Range(Array("Powiat")).Glow.Radius <> 0 Then

        pozname = mapa(row, column, 0)
        num = numer(row, column, 0) + 2
        If UCase(Left(pozname, 1)) = Left(pozname, 1) Then title = "Miasto na prawach powiatu " & pozname & " - prognoza liczby ludnoœci na lata 2011-2035" Else title = "Powiat " & pozname & " - prognoza liczby ludnoœci na lata 2011-2035"
           
    ElseIf ActiveSheet.Shapes.Range(Array("Woj")).Glow.Radius <> 0 Then
    
        pozname = Woj(row, column, 0)
        num = WojNum(row, column, 0) + 407
        title = "Województwo " & pozname & " - prognoza liczby ludnoœci na lata 2011-2035"
        
    ElseIf ActiveSheet.Shapes.Range(Array("Kraj")).Glow.Radius <> 0 Then Exit Function
        
    End If
    
    If num <> 2 And num <> 407 Then
    
        Sheets("Powiaty").Range("komX").Value = num
           
        With ActiveSheet.ChartObjects("Wykres 1").Chart

            .ChartTitle.Text = title

        End With
    
    End If
       
    If pozname = counter And pozname <> "" Then Exit Function

    With ActiveSheet.Shapes("pole tekstowe 2")

        .Visible = False
        .TextFrame.Characters.Text = ""
        .Top = 1000

    End With

    With Sheets("Powiaty").Shapes(counter)

        .Fill.Transparency = 0

    End With
    
    counter = pozname

    If pozname = "" Then Exit Function

    With ActiveSheet.Shapes("pole tekstowe 2")

        .TextFrame.Characters.Text = pozname
        .TextFrame.AutoSize = True
        .Top = ActiveSheet.Cells(row, column).Top + 10
        .Left = ActiveSheet.Cells(row, column).Left + 15
        .Visible = True

    End With

    With Sheets("Powiaty").Shapes(pozname)

        .Fill.Transparency = 0.7

    End With
    
    Sheets("Mapka dane").Calculate
    Range("YAKRES").Calculate
                 
End Function
Public Function Odswiez()

    Dim i, j, num As Integer

    For i = 1 To 57

        For j = 1 To 70

            mapa(i, j, 0) = Sheets("Mapy pomocnicze").Cells(i, j)
            numer(i, j, 0) = Sheets("Mapy pomocnicze").Cells(i + 60, j)
            Woj(i, j, 0) = Sheets("Mapy pomocnicze").Cells(i + 120, j)
            WojNum(i, j, 0) = Sheets("Mapy pomocnicze").Cells(i + 180, j)

        Next j

    Next i
    
End Function
Public Function MapkaPow()

    Dim i, j, x, k, l, m, licz, licznik(1 To 55, 1 To 64), counter1 As Integer
    Dim dystans(0 To 5) As Long
    Dim obsz As Shape

    For Each obsz In Sheets("Powiaty").Shapes

    counter1 = counter1 + 1

        For i = 1 To 64

            For j = 1 To 55

                If (Sheets("Powiaty").Cells(j, i).Left + Sheets("Powiaty").Cells(j, i).Width / 2 >= Sheets("Powiaty").Shapes(obsz.Name).Left) And (Sheets("Powiaty").Cells(j, i).Left + Sheets("Powiaty").Cells(j, i).Width / 2 <= Sheets("Powiaty").Shapes(obsz.Name).Left + Sheets("Powiaty").Shapes(obsz.Name).Width) And (Sheets("Powiaty").Cells(j, i).Top + Sheets("Powiaty").Cells(j, i).Height / 2 >= Sheets("Powiaty").Shapes(obsz.Name).Top) And (Sheets("Powiaty").Cells(j, i).Top + Sheets("Powiaty").Cells(j, i).Height / 2 <= Sheets("Powiaty").Shapes(obsz.Name).Top + Sheets("Powiaty").Shapes(obsz.Name).Height) Then

                    If mapa(j, i, 0) <> "" Then

                        licznik(j, i) = licznik(j, i) + 1
                        mapa(j, i, licznik(j, i)) = obsz.Name
                        GoTo nast

                    Else

                        mapa(j, i, 0) = obsz.Name
                        GoTo nast

                    End If

                End If
nast:
            Next j

        Next i

    Next obsz

    For k = 1 To 64

        For l = 1 To 55

            licz = 0

            If mapa(l, k, 1) <> "" Then

                Do While (mapa(l, k, licz)) <> ""

                    dystans(licz) = Sqr((Sheets("Powiaty").Shapes(mapa(l, k, licz)).Left + Sheets("Powiaty").Shapes(mapa(l, k, licz)).Width / 2 - Sheets("Powiaty").Cells(l, k).Left + Sheets("Powiaty").Cells(l, k).Width / 2) * (Sheets("Powiaty").Shapes(mapa(l, k, licz)).Left + Sheets("Powiaty").Shapes(mapa(l, k, licz)).Width / 2 - Sheets("Powiaty").Cells(l, k).Left + Sheets("Powiaty").Cells(l, k).Width / 2) + (Sheets("Powiaty").Shapes(mapa(l, k, licz)).Top + Sheets("Powiaty").Shapes(mapa(l, k, licz)).Height / 2 - Sheets("Powiaty").Cells(l, k).Top + Sheets("Powiaty").Cells(l, k).Height / 2) * (Sheets("Powiaty").Shapes(mapa(l, k, licz)).Top + Sheets("Powiaty").Shapes(mapa(l, k, licz)).Height / 2 - Sheets("Powiaty").Cells(l, k).Top + Sheets("Powiaty").Cells(l, k).Height / 2))
                    licz = licz + 1

                Loop

                For m = 0 To licz - 1

                    If UCase(Left(mapa(l, k, m), 1)) = Left(mapa(l, k, m), 1) Then

                        mapa(l, k, 0) = mapa(l, k, m)
                        GoTo nast2

                    End If

                    If dystans(m) < dystans(m + 1) Then

                        dystans(m + 1) = dystans(m)
                        mapa(l, k, m + 1) = mapa(l, k, m)

                    End If

                Next m

                mapa(l, k, 0) = mapa(l, k, licz - 1)

            End If
nast2:

        Next l

    Next k

End Function

Public Function MapkaNum()

    Dim i, j, x, k, l, m, licz, licznik(1 To 55, 1 To 64), counter1 As Integer
    Dim dystans(0 To 5) As Long
    Dim obsz As Shape

    For Each obsz In Sheets("Powiaty").Shapes

    counter1 = counter1 + 1

        For i = 1 To 64

            For j = 1 To 55

                If (Sheets("Powiaty").Cells(j, i).Left + Sheets("Powiaty").Cells(j, i).Width / 2 >= Sheets("Powiaty").Shapes(obsz.Name).Left) And (Sheets("Powiaty").Cells(j, i).Left + Sheets("Powiaty").Cells(j, i).Width / 2 <= Sheets("Powiaty").Shapes(obsz.Name).Left + Sheets("Powiaty").Shapes(obsz.Name).Width) And (Sheets("Powiaty").Cells(j, i).Top + Sheets("Powiaty").Cells(j, i).Height / 2 >= Sheets("Powiaty").Shapes(obsz.Name).Top) And (Sheets("Powiaty").Cells(j, i).Top + Sheets("Powiaty").Cells(j, i).Height / 2 <= Sheets("Powiaty").Shapes(obsz.Name).Top + Sheets("Powiaty").Shapes(obsz.Name).Height) Then

                    If numer(j, i, 0) <> 0 Then

                        licznik(j, i) = licznik(j, i) + 1
                        numer(j, i, licznik(j, i)) = counter1
                        GoTo nast

                    Else

                        numer(j, i, 0) = counter1
                        GoTo nast

                    End If

                End If
nast:
            Next j

        Next i

    Next obsz

    For k = 1 To 64

        For l = 1 To 55

            licz = 0

            If numer(l, k, 1) <> 0 Then

                Do While (numer(l, k, licz)) <> 0

                    dystans(licz) = Sqr((Sheets("Powiaty").Shapes(numer(l, k, licz)).Left + Sheets("Powiaty").Shapes(numer(l, k, licz)).Width / 2 - Sheets("Powiaty").Cells(l, k).Left + Sheets("Powiaty").Cells(l, k).Width / 2) * (Sheets("Powiaty").Shapes(numer(l, k, licz)).Left + Sheets("Powiaty").Shapes(numer(l, k, licz)).Width / 2 - Sheets("Powiaty").Cells(l, k).Left + Sheets("Powiaty").Cells(l, k).Width / 2) + (Sheets("Powiaty").Shapes(numer(l, k, licz)).Top + Sheets("Powiaty").Shapes(numer(l, k, licz)).Height / 2 - Sheets("Powiaty").Cells(l, k).Top + Sheets("Powiaty").Cells(l, k).Height / 2) * (Sheets("Powiaty").Shapes(numer(l, k, licz)).Top + Sheets("Powiaty").Shapes(numer(l, k, licz)).Height / 2 - Sheets("Powiaty").Cells(l, k).Top + Sheets("Powiaty").Cells(l, k).Height / 2))
                    licz = licz + 1

                Loop

                For m = 0 To licz - 1

                    If dystans(m) < dystans(m + 1) Then

                        dystans(m + 1) = dystans(m)
                        numer(l, k, m + 1) = numer(l, k, m)

                    End If

                Next m

                numer(l, k, 0) = numer(l, k, licz - 1)

            End If
nast2:

        Next l

    Next k

End Function

Public Function MapkaWoj()

    Dim i, j, x, k, l, m, licz, licznik(1 To 55, 1 To 64) As Integer
    Dim dystans(0 To 5) As Long
    Dim obsz As Shape
    
    For Each obsz In Sheets("Powiaty").Shapes
            
        For i = 1 To 64
        
            For j = 1 To 55
            
                If (Sheets("Powiaty").Cells(j, i).Left + Sheets("Powiaty").Cells(j, i).Width / 2 >= Sheets("Powiaty").Shapes(obsz.Name).Left) And (Sheets("Powiaty").Cells(j, i).Left + Sheets("Powiaty").Cells(j, i).Width / 2 <= Sheets("Powiaty").Shapes(obsz.Name).Left + Sheets("Powiaty").Shapes(obsz.Name).Width) And (Sheets("Powiaty").Cells(j, i).Top + Sheets("Powiaty").Cells(j, i).Height / 2 >= Sheets("Powiaty").Shapes(obsz.Name).Top) And (Sheets("Powiaty").Cells(j, i).Top + Sheets("Powiaty").Cells(j, i).Height / 2 <= Sheets("Powiaty").Shapes(obsz.Name).Top + Sheets("Powiaty").Shapes(obsz.Name).Height) Then
                
                    If Woj(j, i, 0) <> "" Then
                        
                        licznik(j, i) = licznik(j, i) + 1
                        Woj(j, i, licznik(j, i)) = obsz.Name
                        GoTo nast
                    
                    Else
                    
                        Woj(j, i, 0) = obsz.Name
                        GoTo nast

                    End If
                                    
                End If
nast:
            Next j
            
        Next i
    
    Next obsz
    
    For k = 1 To 64
        
        For l = 1 To 55
        
            licz = 0
                    
            If Woj(l, k, 1) <> "" Then
            
                Do While (Woj(l, k, licz)) <> ""
                
                    dystans(licz) = Sqr((Sheets("Powiaty").Shapes(Woj(l, k, licz)).Left + Sheets("Powiaty").Shapes(Woj(l, k, licz)).Width / 2 - Sheets("Powiaty").Cells(l, k).Left + Sheets("Powiaty").Cells(l, k).Width / 2) * (Sheets("Powiaty").Shapes(Woj(l, k, licz)).Left + Sheets("Powiaty").Shapes(Woj(l, k, licz)).Width / 2 - Sheets("Powiaty").Cells(l, k).Left + Sheets("Powiaty").Cells(l, k).Width / 2) + (Sheets("Powiaty").Shapes(Woj(l, k, licz)).Top + Sheets("Powiaty").Shapes(Woj(l, k, licz)).Height / 2 - Sheets("Powiaty").Cells(l, k).Top + Sheets("Powiaty").Cells(l, k).Height / 2) * (Sheets("Powiaty").Shapes(Woj(l, k, licz)).Top + Sheets("Powiaty").Shapes(Woj(l, k, licz)).Height / 2 - Sheets("Powiaty").Cells(l, k).Top + Sheets("Powiaty").Cells(l, k).Height / 2))
                    licz = licz + 1
                    
                Loop
                
                For m = 0 To licz - 1
   
                    If dystans(m) < dystans(m + 1) Then
                        
                        dystans(m + 1) = dystans(m)
                        Woj(l, k, m + 1) = Woj(l, k, m)
                                                
                    End If
        
                Next m
                
                Woj(l, k, 0) = Woj(l, k, licz - 1)
            
            End If
nast2:
        
        Next l
    
    Next k

End Function

Public Function MapkaWojNum()

    Dim i, j, x, k, l, m, licz, licznik(1 To 55, 1 To 64), counter1 As Integer
    Dim dystans(0 To 5) As Long
    Dim obsz As Shape
    
    For Each obsz In Sheets("Powiaty").Shapes
    
        counter1 = counter1 + 1
            
        For i = 1 To 64
        
            For j = 1 To 55
            
                If (Sheets("Powiaty").Cells(j, i).Left + Sheets("Powiaty").Cells(j, i).Width / 2 >= Sheets("Powiaty").Shapes(obsz.Name).Left) And (Sheets("Powiaty").Cells(j, i).Left + Sheets("Powiaty").Cells(j, i).Width / 2 <= Sheets("Powiaty").Shapes(obsz.Name).Left + Sheets("Powiaty").Shapes(obsz.Name).Width) And (Sheets("Powiaty").Cells(j, i).Top + Sheets("Powiaty").Cells(j, i).Height / 2 >= Sheets("Powiaty").Shapes(obsz.Name).Top) And (Sheets("Powiaty").Cells(j, i).Top + Sheets("Powiaty").Cells(j, i).Height / 2 <= Sheets("Powiaty").Shapes(obsz.Name).Top + Sheets("Powiaty").Shapes(obsz.Name).Height) Then
                
                    If WojNum(j, i, 0) <> 0 Then
                        
                        licznik(j, i) = licznik(j, i) + 1
                        WojNum(j, i, licznik(j, i)) = counter1
                        GoTo nast
                    
                    Else
                    
                        WojNum(j, i, 0) = counter1
                        GoTo nast

                    End If
                                    
                End If
nast:
            Next j
            
        Next i
    
    Next obsz
    
    For k = 1 To 64
        
        For l = 1 To 55
        
            licz = 0
                    
            If WojNum(l, k, 1) <> 0 Then
            
                Do While (WojNum(l, k, licz)) <> 0
                
                    dystans(licz) = Sqr((Sheets("Powiaty").Shapes(WojNum(l, k, licz)).Left + Sheets("Powiaty").Shapes(WojNum(l, k, licz)).Width / 2 - Sheets("Powiaty").Cells(l, k).Left + Sheets("Powiaty").Cells(l, k).Width / 2) * (Sheets("Powiaty").Shapes(WojNum(l, k, licz)).Left + Sheets("Powiaty").Shapes(WojNum(l, k, licz)).Width / 2 - Sheets("Powiaty").Cells(l, k).Left + Sheets("Powiaty").Cells(l, k).Width / 2) + (Sheets("Powiaty").Shapes(WojNum(l, k, licz)).Top + Sheets("Powiaty").Shapes(WojNum(l, k, licz)).Height / 2 - Sheets("Powiaty").Cells(l, k).Top + Sheets("Powiaty").Cells(l, k).Height / 2) * (Sheets("Powiaty").Shapes(WojNum(l, k, licz)).Top + Sheets("Powiaty").Shapes(WojNum(l, k, licz)).Height / 2 - Sheets("Powiaty").Cells(l, k).Top + Sheets("Powiaty").Cells(l, k).Height / 2))
                    licz = licz + 1
                    
                Loop
                
                For m = 0 To licz - 1
   
                    If dystans(m) < dystans(m + 1) Then
                        
                        dystans(m + 1) = dystans(m)
                        WojNum(l, k, m + 1) = WojNum(l, k, m)
                                                
                    End If
        
                Next m
                
                WojNum(l, k, 0) = WojNum(l, k, licz - 1)
            
            End If
nast2:
        
        Next l
    
    Next k

End Function

Public Function IE_Czyszczenie() 'wy³¹czanie wszystkich okienek internet explorera (stosuje, bo nie zawsze sie wy³¹czaj¹ bo przwo³aniu)

    Dim objWMI As Object, objProcess As Object, objProcesses As Object
    Set objWMI = GetObject("winmgmts://.")
    Set objProcesses = objWMI.ExecQuery( _
        "SELECT * FROM Win32_Process WHERE Name = 'iexplore.exe'")
    For Each objProcess In objProcesses
    On Error Resume Next
        Call objProcess.Terminate
    On Error GoTo 0
    Next
    Set objProcesses = Nothing: Set objWMI = Nothing
    
End Function

'kod ze strony http://www.cpearson.com/excel/zoom.htm
Public Function ZoomToRange(ByVal ZoomThisRange As Range, ByVal PreserveRows As Boolean)

        Dim Wind As Window
        
        Set Wind = ActiveWindow
        Application.ScreenUpdating = False
        '
        ' Put the upper left cell of the range in the top-left of the screen.
        '
        Application.Goto ZoomThisRange(1, 1), True
        
        With ZoomThisRange
            If PreserveRows = True Then
                .Resize(.Rows.Count, 1).Select
            Else
                .Resize(1, .Columns.Count).Select
            End If
        End With
        
        With Wind
            .zoom = True
            .VisibleRange(1, 1).Select
        End With

End Function
