Attribute VB_Name = "nyelv"
Option Explicit
Public szoveg(1 To 30) As String
Public utasitasok(1 To 20) As String

Public Sub magyar_nyelv()
    With emulator
        .menu(0).ToolTipText = "�j program - [CTRL]+[U]"
        .menu(1).ToolTipText = "Program megnyit�sa - [CTRL]+[O]"
        .menu(2).ToolTipText = "Program ment�se - [CTRL]+[S]"
        .menu(3).ToolTipText = "Sor besz�r�sa - [Insert]"
        .menu(13).ToolTipText = "Program Import�l�sa - [CTRL]+[I]"
        .menu(4).ToolTipText = "Sor m�dos�t�sa - [Enter]"
        .menu(9).ToolTipText = "Mozgat�sa felfel� - [-]"
        .menu(10).ToolTipText = "Mozgat�s lefel� - [+]"
        .menu(5).ToolTipText = "Sor t�rl�se - [Delete]"
        .menu(6).ToolTipText = "Program futtat�sa - [F5]"
        .menu(7).ToolTipText = "Futtat�s l�p�senk�nt - [F8]"
        .menu(8).ToolTipText = "Futtat�s megszak�t�sa - [Break]"
        .menu(11).ToolTipText = "N�vjegy"
        .menu(12).ToolTipText = "Kil�p�s"
                    
        'Rendszer
        .regpanel.Caption = "Rendszer:"
        .cimke(11).Caption = "Regiszterek"
        .cimke(8).Caption = "Aktu�lis utas�t�s:"
        .cimke(9).Caption = "Poz�ci�:"
        .cimke(10).Caption = "Verem:"
        .cimke(13).Caption = "Kimenet:"
        .cimke(12).Caption = "Feldolgoz�s sebess�ge:"
        .sebesseg(0).Caption = "Maxim�lis"
        .sebesseg(1).Caption = "1 m�velet/mp"
        .sebesseg(2).Caption = "1 m�velet/2 mp"
        .sebesseg(3).Caption = "1 m�velet/5 mp"
        .cimke(14).Caption = "Nyelv:"
    End With
    
        'szerkeszt�
        kodablak.Caption = "K�d szerkeszt�"
        kodablak.cimke(0).Caption = "Utas�t�s:"
        kodablak.cimke(1).Caption = "1. param�ter:"
        kodablak.cimke(2).Caption = "2. param�ter:"
        kodablak.cimke(3).Caption = "Megjegyz�s:"
        kodablak.felvesz = "&Felvesz"
        kodablak.modosit = "M�do&s�t"
        kodablak.megse.Caption = "&M�gse"
        
        nevjegy.cimke(1).Caption = "8 regiszteres egyszer�s�tett assembly emulator, amely seg�t elsaj�t�tani �s meg�rteni assembly nyelvet."
        nevjegy.cimke(2).Caption = "Ford�t�s:"
        nevjegy.cimke(0).Caption = "Eredeti magyar nyelv� v�ltozat"
        nevjegy.cimke(3).Caption = "K�sz�net:"
        
        'Sz�vegek:
        szoveg(1) = "K�zvetlen �rt�kad�s ismeretlen regiszterb�l."
        szoveg(2) = "Nem adhat �rt�ket neml�tez� regiszternek!"
        szoveg(3) = "K�rem adja meg a(z) <!> regiszter �j �rt�k�t:"
        szoveg(4) = "Ki�r�s:"
        szoveg(5) = "Ismeretlen utas�t�s"
        szoveg(6) = "A program <!> l�p�sben futott le."
        szoveg(7) = "A mem�ria t�rl�s�hez kattintson a Meg�ll�t gombra!"
        
        szoveg(8) = "Biztos �j programot akar kezdeni?"
        szoveg(9) = "�j program..."
        
        szoveg(10) = "Val�ban t�r�lni akarja a kiv�lasztott sort?"
        szoveg(11) = "T�rl�s meger�s�t�se..."
        
        szoveg(12) = "Biztos ki akar l�pni?"
        szoveg(13) = "Kil�p�s..."
        
        szoveg(14) = "Assembly program megnyit�sa..."
        szoveg(15) = "Sz�veges f�jl"
        szoveg(16) = "Assembly program"
        szoveg(17) = "Minden f�jl"
        
        szoveg(18) = "Assembly program import�l�sa..."
        szoveg(19) = "Assembly program ment�se..."
        
        'Parancs s�g�
        utasitasok(1) = "Megjegyz�s besz�r�sa."
        utasitasok(2) = "Sz�m bek�r�se a felhaszn�l�t�l. INP <reg> <alap�rt�k>"
        utasitasok(3) = "Sz�m beolvas�sa egy mem�riac�mr�l. LET <reg> <c�m>"
        utasitasok(4) = "Sz�m t�rol�sa az adott mem�riac�men. STR <sz�m>"
        utasitasok(5) = "K�t sz�m �sszead�sa. Eredm�ny t�rol�sa az els� regiszterbe. ADD <reg> <reg>"
        utasitasok(6) = "Kivon�s az els� �rt�kb�l. Az eredm�ny az els� regiszterbe ker�l. SUB <reg> <reg>"
        utasitasok(7) = "Regiszter �rt�k�nek n�vel�se egyel. INC <reg>"
        utasitasok(8) = "Regiszter �rt�k�nek cs�kkent�se egyel. DEC <reg>"
        utasitasok(9) = "K�t sz�m szorzata. Az eredm�ny az els� regiszterben t�rol�dik el. MLP <reg> <reg>"
        utasitasok(10) = "Eg�sz oszt�s. Az eredm�ny az els� regiszterben t�rol�dik el. DIV <reg> <reg>"
        utasitasok(11) = "M�sodik regiszter �rt�k�nek bem�sol�sa ez els�be. MOV <reg> <reg>"
        utasitasok(12) = "Ha X=Y, akkor ugr�s a megadott mem�riac�mre. JMP <c�m>"
        utasitasok(13) = "Ha X>Y, akkor ugr�s a megadott mem�riac�mre. SIG <c�m>"
        utasitasok(14) = "Ugr�s a megadott mem�riac�mre. GTO <c�m>"
        utasitasok(15) = "Elj�r�sh�v�s a megadott mem�riac�mt�l. RET-tel visszat�r�s a k�vetkez� sorhoz. GSB <c�m>"
        utasitasok(16) = "Visszaugr�s az elj�r�sh�v�si pontot k�vet� mem�riac�mre. RET <c�m>"
        utasitasok(17) = "Regiszter vagy sz�m �rt�k�nek ki�rat�sa a felhaszn�l�nak. OUT <�rt�k>"
        utasitasok(18) = "A program fut�s�nak megszak�t�sa."
        
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
                    'Men�
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
                    
                    'szerkeszt�
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
