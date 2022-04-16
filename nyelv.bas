Attribute VB_Name = "nyelv"
Option Explicit
Public szoveg(1 To 30) As String
Public utasitasok(1 To 20) As String

Public Sub magyar_nyelv()
    With emulator
        .menu(0).ToolTipText = "Új program - [CTRL]+[U]"
        .menu(1).ToolTipText = "Program megnyitása - [CTRL]+[O]"
        .menu(2).ToolTipText = "Program mentése - [CTRL]+[S]"
        .menu(3).ToolTipText = "Sor beszúrása - [Insert]"
        .menu(13).ToolTipText = "Program Importálása - [CTRL]+[I]"
        .menu(4).ToolTipText = "Sor módosítása - [Enter]"
        .menu(9).ToolTipText = "Mozgatása felfelé - [-]"
        .menu(10).ToolTipText = "Mozgatás lefelé - [+]"
        .menu(5).ToolTipText = "Sor törlése - [Delete]"
        .menu(6).ToolTipText = "Program futtatása - [F5]"
        .menu(7).ToolTipText = "Futtatás lépésenként - [F8]"
        .menu(8).ToolTipText = "Futtatás megszakítása - [Break]"
        .menu(11).ToolTipText = "Névjegy"
        .menu(12).ToolTipText = "Kilépés"
                    
        'Rendszer
        .regpanel.Caption = "Rendszer:"
        .cimke(11).Caption = "Regiszterek"
        .cimke(8).Caption = "Aktuális utasítás:"
        .cimke(9).Caption = "Pozíció:"
        .cimke(10).Caption = "Verem:"
        .cimke(13).Caption = "Kimenet:"
        .cimke(12).Caption = "Feldolgozás sebessége:"
        .sebesseg(0).Caption = "Maximális"
        .sebesseg(1).Caption = "1 mûvelet/mp"
        .sebesseg(2).Caption = "1 mûvelet/2 mp"
        .sebesseg(3).Caption = "1 mûvelet/5 mp"
        .cimke(14).Caption = "Nyelv:"
    End With
    
        'szerkesztõ
        kodablak.Caption = "Kód szerkesztõ"
        kodablak.cimke(0).Caption = "Utasítás:"
        kodablak.cimke(1).Caption = "1. paraméter:"
        kodablak.cimke(2).Caption = "2. paraméter:"
        kodablak.cimke(3).Caption = "Megjegyzés:"
        kodablak.felvesz = "&Felvesz"
        kodablak.modosit = "Módo&sít"
        kodablak.megse.Caption = "&Mégse"
        
        nevjegy.cimke(1).Caption = "8 regiszteres egyszerûsített assembly emulator, amely segít elsajátítani és megérteni assembly nyelvet."
        nevjegy.cimke(2).Caption = "Fordítás:"
        nevjegy.cimke(0).Caption = "Eredeti magyar nyelvû változat"
        nevjegy.cimke(3).Caption = "Köszönet:"
        
        'Szövegek:
        szoveg(1) = "Közvetlen értékadás ismeretlen regiszterbõl."
        szoveg(2) = "Nem adhat értéket nemlétezõ regiszternek!"
        szoveg(3) = "Kérem adja meg a(z) <!> regiszter új értékét:"
        szoveg(4) = "Kiírás:"
        szoveg(5) = "Ismeretlen utasítás"
        szoveg(6) = "A program <!> lépésben futott le."
        szoveg(7) = "A memória törléséhez kattintson a Megállít gombra!"
        
        szoveg(8) = "Biztos új programot akar kezdeni?"
        szoveg(9) = "Új program..."
        
        szoveg(10) = "Valóban törölni akarja a kiválasztott sort?"
        szoveg(11) = "Törlés megerõsítése..."
        
        szoveg(12) = "Biztos ki akar lépni?"
        szoveg(13) = "Kilépés..."
        
        szoveg(14) = "Assembly program megnyitása..."
        szoveg(15) = "Szöveges fájl"
        szoveg(16) = "Assembly program"
        szoveg(17) = "Minden fájl"
        
        szoveg(18) = "Assembly program importálása..."
        szoveg(19) = "Assembly program mentése..."
        
        'Parancs súgó
        utasitasok(1) = "Megjegyzés beszúrása."
        utasitasok(2) = "Szám bekérése a felhasználótól. INP <reg> <alapérték>"
        utasitasok(3) = "Szám beolvasása egy memóriacímrõl. LET <reg> <cím>"
        utasitasok(4) = "Szám tárolása az adott memóriacímen. STR <szám>"
        utasitasok(5) = "Két szám összeadása. Eredmény tárolása az elsõ regiszterbe. ADD <reg> <reg>"
        utasitasok(6) = "Kivonás az elsõ értékbõl. Az eredmény az elsõ regiszterbe kerül. SUB <reg> <reg>"
        utasitasok(7) = "Regiszter értékének növelése egyel. INC <reg>"
        utasitasok(8) = "Regiszter értékének csökkentése egyel. DEC <reg>"
        utasitasok(9) = "Két szám szorzata. Az eredmény az elsõ regiszterben tárolódik el. MLP <reg> <reg>"
        utasitasok(10) = "Egész osztás. Az eredmény az elsõ regiszterben tárolódik el. DIV <reg> <reg>"
        utasitasok(11) = "Második regiszter értékének bemásolása ez elsõbe. MOV <reg> <reg>"
        utasitasok(12) = "Ha X=Y, akkor ugrás a megadott memóriacímre. JMP <cím>"
        utasitasok(13) = "Ha X>Y, akkor ugrás a megadott memóriacímre. SIG <cím>"
        utasitasok(14) = "Ugrás a megadott memóriacímre. GTO <cím>"
        utasitasok(15) = "Eljáráshívás a megadott memóriacímtõl. RET-tel visszatérés a következõ sorhoz. GSB <cím>"
        utasitasok(16) = "Visszaugrás az eljáráshívási pontot követõ memóriacímre. RET <cím>"
        utasitasok(17) = "Regiszter vagy szám értékének kiíratása a felhasználónak. OUT <érték>"
        utasitasok(18) = "A program futásának megszakítása."
        
        NyelvFrissites
End Sub
Public Sub nyelv(FajlNev As String)
Dim sor As String, parancs As String, Index As Byte
On Error GoTo hiba:
    Open FajlNev For Input As 4
        Do While Not EOF(4)
                Line Input #4, sor
                
                If sor = "" Or Mid(sor, 1, 1) = ";" Or Mid(sor, 1, 1) = "#" Or Mid(sor, 1, 1) = "/" Or Mid(sor, 1, 1) = "[" Then GoTo kihagy
                parancs = Utasitas(sor)
                
                With emulator
                Select Case parancs
                    'Menü
                    Case "fajl.uj"
                        .menu(0).ToolTipText = Ertek(sor) & " - [CTRL]+[U]"
                    Case "fajl.megnyit"
                        .menu(1).ToolTipText = Ertek(sor) & " - [CTRL]+[O]"
                    Case "fajl.ment"
                        .menu(2).ToolTipText = Ertek(sor) & " - [CTRL]+[S]"
                    Case "szerk.beszur"
                        .menu(3).ToolTipText = Ertek(sor) & " - [Insert]"
                    Case "szerk.import"
                        .menu(13).ToolTipText = Ertek(sor) & " - [CTRL]+[I]"
                    Case "szer.modosit"
                        .menu(4).ToolTipText = Ertek(sor) & " - [Enter]"
                    Case "szerk.fel"
                        .menu(9).ToolTipText = Ertek(sor) & " - [-]"
                    Case "szerk.le"
                        .menu(10).ToolTipText = Ertek(sor) & " - [+]"
                    Case "szerk.torol"
                        .menu(5).ToolTipText = Ertek(sor) & " - [Delete]"
                    Case "fut.indit"
                        .menu(6).ToolTipText = Ertek(sor) & " - [F5]"
                    Case "fut.leptet"
                        .menu(7).ToolTipText = Ertek(sor) & " - [F8]"
                    Case "fut.stop"
                        .menu(8).ToolTipText = Ertek(sor) & " - [Break]"
                    Case "fajl.nevjegy"
                        .menu(11).ToolTipText = Ertek(sor)
                    Case "fajl.kilep"
                        .menu(12).ToolTipText = Ertek(sor)
                    
                    'Rendszer
                    Case "rendszer"
                        .regpanel.Caption = Ertek(sor)
                    Case "rendszer.regiszter"
                        .cimke(11).Caption = Ertek(sor)
                    Case "rendszer.utasitas"
                        .cimke(8).Caption = Ertek(sor)
                    Case "rendszer.jelenlegi"
                        .cimke(9).Caption = Ertek(sor)
                    Case "rendszer.verem"
                        .cimke(10).Caption = Ertek(sor)
                    Case "rendszer.kimenet"
                        .cimke(13).Caption = Ertek(sor)
                    Case "rendszer.sebesseg"
                        .cimke(12).Caption = Ertek(sor)
                    Case "rendszer.sebesseg.max"
                        .sebesseg(0).Caption = Ertek(sor)
                    Case "rendszer.sebesseg.1"
                        .sebesseg(1).Caption = Ertek(sor)
                    Case "rendszer.sebesseg.1/2"
                        .sebesseg(2).Caption = Ertek(sor)
                    Case "rendszer.sebesseg.1/5"
                        .sebesseg(3).Caption = Ertek(sor)
                    Case "rendszer.nyelv"
                        .cimke(14).Caption = Ertek(sor)
                    
                    'szerkesztõ
                    Case "kodszerk"
                        kodablak.Caption = Ertek(sor)
                    Case "kodszerk.utasitas"
                        kodablak.cimke(0).Caption = Ertek(sor)
                    Case "kodszerk.p1"
                        kodablak.cimke(1).Caption = Ertek(sor)
                    Case "kodszerk.p2"
                        kodablak.cimke(2).Caption = Ertek(sor)
                    Case "kodszerk.megjegyzes"
                        kodablak.cimke(3).Caption = Ertek(sor)
                    Case "kodszerk.felvesz"
                        kodablak.felvesz = Ertek(sor)
                    Case "kodszerk.modosit"
                        kodablak.modosit = Ertek(sor)
                    Case "kodszerk.megse"
                        kodablak.megse.Caption = Ertek(sor)
                    
                    Case "nevjegy.szoveg"
                        nevjegy.cimke(1).Caption = Ertek(sor)
                    Case "nevjegy.forditas"
                        nevjegy.cimke(2).Caption = Ertek(sor)
                    Case "nevjegy.forditas.szoveg"
                        nevjegy.cimke(0).Caption = Ertek(sor)
                    Case "nevjegy.koszonet"
                        nevjegy.cimke(3).Caption = Ertek(sor)
                        
                    'parancssugo
                    Case "REM"
                        utasitasok(1) = Ertek(sor)
                    Case "INP"
                        utasitasok(2) = Ertek(sor)
                    Case "LET"
                        utasitasok(3) = Ertek(sor)
                    Case "STR"
                        utasitasok(4) = Ertek(sor)
                    Case "ADD"
                        utasitasok(5) = Ertek(sor)
                    Case "SUB"
                        utasitasok(6) = Ertek(sor)
                    Case "INC"
                        utasitasok(7) = Ertek(sor)
                    Case "DEC"
                        utasitasok(8) = Ertek(sor)
                    Case "MLP"
                        utasitasok(9) = Ertek(sor)
                    Case "DIV"
                        utasitasok(10) = Ertek(sor)
                    Case "MOV"
                        utasitasok(11) = Ertek(sor)
                    Case "JMP"
                        utasitasok(12) = Ertek(sor)
                    Case "SIG"
                        utasitasok(13) = Ertek(sor)
                    Case "GTO"
                        utasitasok(14) = Ertek(sor)
                    Case "GSB"
                        utasitasok(15) = Ertek(sor)
                    Case "RET"
                        utasitasok(16) = Ertek(sor)
                    Case "OUT"
                        utasitasok(17) = Ertek(sor)
                    Case "END"
                        utasitasok(18) = Ertek(sor)
                End Select
                
                If IsNumeric(parancs) Then
                    Index = CByte(parancs)
                    szoveg(Index) = Ertek(sor)
                End If
                
                End With


kihagy:
        Loop
    Close 4
    NyelvFrissites
Exit Sub
hiba:
    MsgBox Err.Description
    Close 4
    magyar_nyelv
End Sub
Public Function Utasitas(Adatsor As String) As String
    Dim i As Integer, megvan As Boolean
    i = 1
    megvan = False
    Do While i <= Len(Adatsor) And Not megvan
        If Mid(Adatsor, i, 1) = "=" Then
                    megvan = True
                    Utasitas = Mid(Adatsor, 1, i - 1)
        End If
        i = i + 1
    Loop
    If Not megvan Then Utasitas = Adatsor
End Function
Public Function Ertek(Adatsor As String) As String
    Dim i As Integer, megvan As Boolean
    i = 1
    megvan = False
    Do While i <= Len(Adatsor) And Not megvan
        If Mid(Adatsor, i, 1) = "=" Then
                    megvan = True
                    Ertek = Mid(Adatsor, i + 1, Len(Adatsor) - i)
        End If
        i = i + 1
    Loop
    If Not megvan Then Ertek = ""
End Function
Public Function Szovegbe(Mibe As String, Hova As String, Mit As String) As String
    Dim i As Long
    i = 1
    Szovegbe = ""
    
    Do While i <= Len(Mibe)
        If Mid(Mibe, i, Len(Hova)) = Hova Then
                Szovegbe = Szovegbe & Mit
                i = i + Len(Hova) - 1
            Else
                Szovegbe = Szovegbe & Mid(Mibe, i, 1)
        End If
        i = i + 1
    Loop
End Function

Public Sub NyelvFrissites()
    kodablak.kod_Click
End Sub
