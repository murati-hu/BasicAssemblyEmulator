VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form emulator 
   Caption         =   "BAsE8"
   ClientHeight    =   7980
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9240
   Icon            =   "emulator8.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7980
   ScaleWidth      =   9240
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox nyelvvalaszto 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   49
      Top             =   6840
      Width           =   1935
   End
   Begin VB.FileListBox nyelvmappa 
      Height          =   1260
      Left            =   6960
      Pattern         =   "*.lng"
      TabIndex        =   48
      Top             =   840
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox loglista 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   2280
      TabIndex        =   47
      Top             =   5040
      Width           =   6495
   End
   Begin VB.Frame regpanel 
      Appearance      =   0  'Flat
      Caption         =   "Rendszer:"
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   0
      TabIndex        =   15
      Top             =   600
      Width           =   2175
      Begin VB.TextBox kimenet 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   45
         TabStop         =   0   'False
         Top             =   4080
         Width           =   1215
      End
      Begin VB.ListBox veremlista 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   840
         TabIndex        =   43
         TabStop         =   0   'False
         Top             =   3000
         Width           =   1215
      End
      Begin VB.OptionButton sebesseg 
         Caption         =   "1 mûvelet/ 5 mp"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   5520
         Width           =   1935
      End
      Begin VB.OptionButton sebesseg 
         Caption         =   "1 mûvelet/ 2 mp"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   5280
         Width           =   1935
      End
      Begin VB.OptionButton sebesseg 
         Caption         =   "1 mûvelet/mp"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   5040
         Width           =   1935
      End
      Begin VB.OptionButton sebesseg 
         Caption         =   "Maximális"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   4800
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.TextBox reg 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox reg 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   480
         Width           =   615
      End
      Begin VB.TextBox reg 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox reg 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   840
         Width           =   615
      End
      Begin VB.TextBox reg 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox reg 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1200
         Width           =   615
      End
      Begin VB.TextBox reg 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   360
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox reg 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   1560
         Width           =   615
      End
      Begin VB.TextBox akt_muv 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1935
      End
      Begin VB.TextBox j_cim 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Nyelv:"
         Height          =   195
         Index           =   14
         Left            =   120
         TabIndex        =   50
         Top             =   6000
         Width           =   450
      End
      Begin VB.Line Line5 
         X1              =   0
         X2              =   2160
         Y1              =   5880
         Y2              =   5880
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Kimenet:"
         Height          =   195
         Index           =   13
         Left            =   120
         TabIndex        =   44
         Top             =   4080
         Width           =   660
         WordWrap        =   -1  'True
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Feldolgozás sebessége:"
         Height          =   195
         Index           =   12
         Left            =   120
         TabIndex        =   38
         Top             =   4560
         Width           =   1695
      End
      Begin VB.Line Line8 
         X1              =   0
         X2              =   2160
         Y1              =   4440
         Y2              =   4440
      End
      Begin VB.Line Line7 
         X1              =   0
         X2              =   2160
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Regiszterek"
         Height          =   195
         Index           =   11
         Left            =   600
         TabIndex        =   37
         Top             =   240
         Width           =   840
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "X"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   36
         Top             =   480
         Width           =   105
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "A"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   35
         Top             =   840
         Width           =   105
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "C"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   34
         Top             =   1200
         Width           =   105
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "E"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   33
         Top             =   1560
         Width           =   105
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Y"
         Height          =   195
         Index           =   4
         Left            =   1920
         TabIndex        =   32
         Top             =   480
         Width           =   105
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "B"
         Height          =   195
         Index           =   5
         Left            =   1920
         TabIndex        =   31
         Top             =   840
         Width           =   105
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "D"
         Height          =   195
         Index           =   6
         Left            =   1920
         TabIndex        =   30
         Top             =   1200
         Width           =   120
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "F"
         Height          =   195
         Index           =   7
         Left            =   1920
         TabIndex        =   29
         Top             =   1560
         Width           =   90
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Aktuális utasítás:"
         Height          =   195
         Index           =   8
         Left            =   120
         TabIndex        =   28
         Top             =   2040
         Width           =   1200
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Jelenlegi:"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   27
         Top             =   2640
         Width           =   660
      End
      Begin VB.Label cimke 
         AutoSize        =   -1  'True
         Caption         =   "Verem:"
         Height          =   195
         Index           =   10
         Left            =   120
         TabIndex        =   26
         Top             =   3000
         Width           =   615
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox vezerlo_tarto 
      Align           =   1  'Align Top
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   0
      ScaleHeight     =   615
      ScaleWidth      =   9240
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   0
      Width           =   9240
      Begin VB.PictureBox menu 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   13
         Left            =   2520
         Picture         =   "emulator8.frx":0CCA
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   46
         TabStop         =   0   'False
         ToolTipText     =   "Program importálása [CTRL] + [I]"
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox menu 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   0
         Left            =   0
         Picture         =   "emulator8.frx":1C3C
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   14
         TabStop         =   0   'False
         ToolTipText     =   "Új program - [CTRL]+[U]"
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox menu 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   1
         Left            =   600
         Picture         =   "emulator8.frx":2BAE
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   13
         TabStop         =   0   'False
         ToolTipText     =   "Program megnyitása - [CTRL] + [O]"
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox menu 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   2
         Left            =   1200
         Picture         =   "emulator8.frx":3B20
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   12
         TabStop         =   0   'False
         ToolTipText     =   "Program mentése - [CTRL] + [S]"
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox menu 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   3
         Left            =   1920
         Picture         =   "emulator8.frx":4A92
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Sor beszúrása [Insert]"
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox menu 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   4
         Left            =   3120
         Picture         =   "emulator8.frx":5A04
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   10
         TabStop         =   0   'False
         ToolTipText     =   "Sor módosítása [Enter]"
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox menu 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   5
         Left            =   4200
         Picture         =   "emulator8.frx":6976
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Sor törlése [Delete]"
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox menu 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   6
         Left            =   4920
         Picture         =   "emulator8.frx":78E8
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Futtatás... [F5]"
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox menu 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   7
         Left            =   5520
         Picture         =   "emulator8.frx":852A
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Futtatás lépésenként [F8]"
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox menu 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   8
         Left            =   6120
         Picture         =   "emulator8.frx":916C
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Megállítás [Break]"
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox menu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   9
         Left            =   3840
         Picture         =   "emulator8.frx":9DAE
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Mozgatás felfelé [-]"
         Top             =   0
         Width           =   270
      End
      Begin VB.PictureBox menu 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   270
         Index           =   10
         Left            =   3840
         Picture         =   "emulator8.frx":A1E0
         ScaleHeight     =   270
         ScaleWidth      =   270
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Mozgatás lefelé [+]"
         Top             =   240
         Width           =   270
      End
      Begin VB.PictureBox menu 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   11
         Left            =   6840
         Picture         =   "emulator8.frx":A612
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Súgó"
         Top             =   0
         Width           =   495
      End
      Begin VB.PictureBox menu 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   495
         Index           =   12
         Left            =   7440
         Picture         =   "emulator8.frx":B584
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Kilépés"
         Top             =   0
         Width           =   495
      End
      Begin VB.Line Line1 
         X1              =   1800
         X2              =   1800
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Line Line2 
         X1              =   4800
         X2              =   4800
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   8040
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Line Line4 
         X1              =   6720
         X2              =   6720
         Y1              =   0
         Y2              =   600
      End
      Begin VB.Line Line6 
         X1              =   8040
         X2              =   8040
         Y1              =   0
         Y2              =   600
      End
   End
   Begin MSComDlg.CommonDialog pb 
      Left            =   6000
      Top             =   4920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.Timer idozito 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4680
      Top             =   4680
   End
   Begin VB.ListBox code 
      Appearance      =   0  'Flat
      Columns         =   2
      Height          =   3345
      Left            =   2280
      TabIndex        =   0
      Top             =   720
      Width           =   4575
   End
End
Attribute VB_Name = "emulator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const nevem = "Basic Assemby Emulator 8(BAsE8) v0.1 - Muráti Ákos"

Dim X, Y, a, b, c, d, e, f
Dim Fut As Boolean
'memóriacímek kezelése
Dim aktualis_v As Long, visszater As Long
'Egyéb
Dim UccsoGomb As Byte
Dim lepes As Long, kimeno As String

Private Function Regiszterbol(Melyikbol As String)
    Melyikbol = UCase(Melyikbol)
    Select Case Melyikbol
        Case "X"
            Regiszterbol = X
        Case "Y"
            Regiszterbol = Y
        Case "A"
            Regiszterbol = a
        Case "B"
            Regiszterbol = b
        Case "C"
            Regiszterbol = c
        Case "D"
            Regiszterbol = d
        Case "E"
            Regiszterbol = e
        Case "F"
            Regiszterbol = f
        Case Else
            Regiszterbol = Melyikbol
            Loggol szoveg(1)
    End Select
End Function
Private Sub Regiszterbe(Melyikbe As String, Mit)
    Melyikbe = UCase(Melyikbe)
    Select Case Melyikbe
        Case "X"
            X = Mit
        Case "Y"
            Y = Mit
        Case "A"
            a = Mit
        Case "B"
            b = Mit
        Case "C"
            c = Mit
        Case "D"
            d = Mit
        Case "E"
            e = Mit
        Case "F"
           f = Mit
        Case Else
            Loggol szoveg(2)
    End Select
    Frissit
End Sub
Private Sub Vegrehajt(SorSzam As Long, Utasitas As String, p1 As String, p2 As String)
Dim s1, s2

On Error GoTo hiba
    
    'If p1 <> "" Then s1 = Regiszterbol(p1)
    'If p2 <> "" Then s2 = Regiszterbol(p2)
    
    Utasitas = UCase(Utasitas)
    
    Select Case Utasitas
        Case "ADD" 'Összeadás
            s1 = Regiszterbol(p1)
            s2 = Regiszterbol(p2)
            
            s1 = Szamma(s1) + Szamma(s2)
            Regiszterbe p1, s1
        
        Case "SUB" 'Kivonás
            s1 = Regiszterbol(p1)
            s2 = Regiszterbol(p2)
            
            s1 = Szamma(s1) - Szamma(s2)
            Regiszterbe p1, s1

        Case "MLP" 'Szorzás
            s1 = Regiszterbol(p1)
            s2 = Regiszterbol(p2)
            
            s1 = Szamma(s1) * Szamma(s2)
            Regiszterbe p1, s1
            
        Case "DIV" 'Egészosztás
            s1 = Regiszterbol(p1)
            s2 = Regiszterbol(p2)
            
            s1 = Szamma(s1) \ Szamma(s2) 'Egész osztás
            's1 = CLng(s1)
            Regiszterbe p1, s1
            
        Case "INC" 'Növelés 1-el
            s1 = Regiszterbol(p1)
            
            s1 = Szamma(s1) + 1
            Regiszterbe p1, s1
            
        Case "DEC" 'Csökkentés egyel
            s1 = Regiszterbol(p1)

            s1 = Szamma(s1) - 1
            Regiszterbe p1, s1
            
        Case "MOV" 'Mozgatás
            s1 = Regiszterbol(p1)
            s2 = Regiszterbol(p2)
            
            s1 = s2
            Regiszterbe p1, s1
            
        Case "INP" 'Beolvasás
            's1 = Regiszterbol(p1)
            
            If p2 = "" Then
                    s1 = Szamma(InputBox(Szovegbe(szoveg(3), "<!>", p1), "INP"))
                Else
                    s1 = p2
            End If
            'If s1 = "" Then s1 = 0
            Regiszterbe p1, s1
            
        Case "OUT" 'Kiírás
            s1 = Regiszterbol(p1)
            
            'MsgBox s1, vbOKOnly, "A(z) " & p1 & " regiszter értéke:"
            Loggol szoveg(4) & " " & p1 & "=" & s1
            kimeno = s1
            
        Case "LET" 'Érték kiolvasása
            s1 = Regiszterbol(p1)
            s2 = Regiszterbol(p2)
            
            Regiszterbe p1, Parameter(code.List(s2 - 1), 1)
            
        Case "GTO" 'Ugrás
            's1 = Regiszterbol(p1)
            
            aktualis = p1 - 2 's1 - 2
        
        Case "GSB" 'Eljárás hívása
             's1 = Regiszterbol(p1)
             
             Verem = aktualis + 1
             aktualis = p1 - 2 's1 - 2
             
        Case "RET" 'Eljárás visszatérése
            aktualis = Verem - 1
        
        Case "SIG" 'Ellenõrzés
            's1 = Regiszterbol(p1)
            
            If Szamma(X) > Szamma(Y) Then aktualis = p1 - 2 's1 - 2
        
        Case "END" 'Program vége
            aktualis = code.ListCount '- 1
            
        Case "JMP" 'Összehasonlítás
            's1 = Regiszterbol(p1)
            
            If Szamma(X) = Szamma(Y) Then aktualis = p1 - 2 's1 - 2
        Case "REM"
            
        Case Else
            Loggol szoveg(5)
    End Select
    Frissit
Exit Sub
hiba:
    MsgBox Err.Description, vbInformation, Err.Number
    Loggol Err.Description, Err.Number
    Resume Next
End Sub

Private Sub cimke_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub code_DblClick()
    If code.ListCount > 0 Then kodablak.Mutasd code.ListIndex, True
End Sub


Private Sub code_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Frissit
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'Futtatási billentyû
    Select Case KeyCode
        Case 13
            If Fut Then menu_Click (7)
        Case 116
            menu_Click (6)
        Case 119
            menu_Click (7)
        Case 121, 123, 19, 27
            menu_Click (8)
    End Select
    
If Fut Then Exit Sub
    
    Select Case Shift
        Case 0
            Select Case KeyCode
                Case 45 'insert
                    menu_Click (3)
                Case 46 'del
                    menu_Click (5)
                Case 13 'Enter
                    code_DblClick
                Case 109 '- fel
                    menu_Click (9)
                Case 107 '+ le
                    menu_Click (10)
            End Select
            
        Case 1 'SHIFT
            Select Case KeyCode
                Case 13 'Enter
                    menu_Click (3)
            End Select
            
        Case 2 'CTRL
            Select Case KeyCode
                Case 79 'Megnyitás - O
                    menu_Click (1)
                Case 73 'Iportálás
                    menu_Click (13)
                Case 85 'Új - U
                    menu_Click (0)
                Case 83 'Mentés -S,
                    menu_Click (2)
            End Select
        Case 4
            'MsgBox KeyCode
    End Select
    'MsgBox KeyCode
    'MsgBox Shift
End Sub



Private Sub Form_Load()
Dim i As Integer
    Nullaz
    Me.Caption = nevem
On Error Resume Next
    nyelvmappa.Path = App.Path & "\nyelvek"
    
    nyelvvalaszto.AddItem "Magyar"
    For i = 0 To nyelvmappa.ListCount - 1
        nyelvvalaszto.AddItem Mid(nyelvmappa.List(i), 1, Len(nyelvmappa.List(i)) - 4)
    Next i
    nyelvvalaszto.Text = nyelvvalaszto.List(0)
    'nyelv.magyar_nyelv
    'Frissit
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Frissit
    UccsoGomb = menu.Count
End Sub

Private Sub Form_Resize()
On Error Resume Next
    If Me.Width < 8190 Then Me.Width = 8190
    If Me.Height < 5565 Then Me.Height = 5565
    
    code.Width = Me.ScaleWidth - code.Left - 100
    code.Columns = code.Width \ 3400
    loglista.Move code.Left, Me.ScaleHeight - loglista.Height - 100, code.Width
    code.Height = loglista.Top - code.Top
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload nevjegy
    Unload kodablak
    End
End Sub

Private Sub idozito_Timer()
'On Error Resume Next
    If aktualis > code.ListCount - 1 Then
        'idozito.Enabled = False
        'aktualis = 0
        'visszater = 0
        'Fut = False
        'Nullaz
        'Frissit
        menu_Click (8)
        Loggol Szovegbe(szoveg(6), "<!>", CStr(lepes)), True
        Loggol "-------------------------", True
        Loggol szoveg(7), True
        Exit Sub
    End If
    lepes = lepes + 1
    futtat (aktualis)
    'code.Selected(aktualis) = True
    aktualis = aktualis + 1
    'code.Selected(aktualis) = True
    Frissit
    'If aktualis > code.ListCount - 1 Then idozito.Enabled = False
End Sub

Private Sub menu_Click(Index As Integer)
Dim seged As String, i As Long

If Not menu(Index).Enabled Then Exit Sub
Select Case Index
        Case 0 'Új program
            If MsgBox(szoveg(8), vbQuestion + vbYesNo, szoveg(9)) = vbYes Then
                code.Clear
                Nullaz
                Me.Caption = nevem
            End If

        Case 1 'Program betöltése
            megnyitas_dlg

        Case 2 'Program mentése
            'If Not menu(2).Enabled Then Exit Sub
            mentes_dlg

        Case 3 'Új sor
            kodablak.Mutasd code.ListIndex + 1, False

        Case 4 'Sor módosítása
            code_DblClick

        Case 5 'Sor törlése
            'If Not menu(5).Enabled Then Exit Sub
            'On Error Resume Next 'Legelsõ elem törlésénél
            'code.Selected(code.ListIndex - 1) = True
            If MsgBox(szoveg(10), vbQuestion + vbYesNo, szoveg(11)) = vbNo Then Exit Sub
            
            i = code.ListIndex
            code.RemoveItem (code.ListIndex) ' + 1)
            
            If i = 0 Then
                    If code.ListCount = 0 Then i = -1
                Else
                    If i = code.ListCount Then i = i - 1
            End If
            If i > -1 Then code.Selected(i) = True
            
            Ujrasorszamoz
            'If code.ListIndex > 0 Then
                ''code.Selected(code.ListIndex) = True
                
                'For i = code.ListIndex - 1 To code.ListCount - 1
                '    SorszamJavitasa i
                'Next i
            'End If
            
        Case 6 'Start
            If Not Fut Then
                    Fut = True
                    'proba
                        aktualis = 0
                        visszater = 0
                    'code.Selected(0) = True
                    Nullaz
            End If
            idozito.Enabled = True
            
        Case 7 'Következõ
            If Fut Then
                    idozito_Timer
                Else
                    menu_Click (8)
                    Fut = True
            End If
            Frissit
            
        Case 8 'Stop
            If Fut Then
                    idozito.Enabled = False
                    Fut = False
                    'Nullaz
                    'Nullaz helyett:
                        aktualis = 0
                        visszater = 0
                Else
                    Nullaz
            End If
            'code.Selected(0) = True
            Frissit
            
        Case 9 'fel
            'If Not menu(9).Enabled Then Exit Sub
            
            'On Error Resume Next
            seged = code.List(code.ListIndex - 1)
            code.List(code.ListIndex - 1) = code.List(code.ListIndex)
            code.List(code.ListIndex) = seged
            code.Selected(code.ListIndex - 1) = True
            
            SorszamJavitasa code.ListIndex
            SorszamJavitasa code.ListIndex + 1
        Case 10 'le
            'If Not menu(10).Enabled Then Exit Sub
            
            'On Error Resume Next
            seged = code.List(code.ListIndex + 1)
            code.List(code.ListIndex + 1) = code.List(code.ListIndex)
            code.List(code.ListIndex) = seged
            code.Selected(code.ListIndex + 1) = True
            
            SorszamJavitasa code.ListIndex
            SorszamJavitasa code.ListIndex - 1
            
        Case 11 'Sugo
            nevjegy.Caption = nevem
            nevjegy.Show vbModal
            
        Case 12 'Kilépés
            If MsgBox(szoveg(12), vbQuestion + vbYesNo, szoveg(13)) = vbYes Then
                Unload Me
            End If
        Case 13 'Program importálása
            import_dlg
    End Select
    Frissit
    If code.Enabled Then code.SetFocus
End Sub

Private Sub menu_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If UccsoGomb <> Index Then
        Frissit
        'menu(Index).Top = 50
        menu(Index).PaintPicture menu(Index).Picture, 10, 30
        UccsoGomb = Index
    End If
End Sub
Private Sub Alaphelyzet()
    Frissit
End Sub
Public Function Utasitas(kod As String) As String
    Utasitas = Darabol(kod, " ", 1)
End Function
Public Function Parameter(kod As String, Melyiket As Byte) As Variant
    Parameter = Darabol(kod, " ", Melyiket + 1)
End Function
Private Function Darabol(szoveg As String, Elvalaszto As String, Melyiket As Byte) As String
    Dim i As Integer, j As Integer, ker As Integer
    Dim cella(0 To 1024) As String
    i = 1
    j = 0
        
    For ker = 1 To Len(szoveg)
        If Mid(szoveg, ker, 1) = Elvalaszto Then
            cella(j) = Mid(szoveg, i, ker - i)
            i = ker + 1
            j = j + 1
        End If
    Next
    cella(j) = Mid(szoveg, i, Len(szoveg) + 1 - i)
    Darabol = cella(Melyiket)
End Function

Private Sub futtat(sor As Long)
    Vegrehajt sor, Utasitas(code.List(sor)), Parameter(code.List(sor), 1), Parameter(code.List(sor), 2)
End Sub
Private Sub Nullaz()
    'Mutatók nullázása
    aktualis = 0
    'visszater = 0
    lepes = 0
    veremlista.Clear
    loglista.Clear
    
    'Regiszterek nullázása
            Regiszterbe "X", 0
            Regiszterbe "Y", 0
            
            Dim i As Integer
            For i = 65 To 70
                Regiszterbe Chr(i), 0
            Next i
End Sub
Private Function Sorszamoz(sor As String, Hanyas As Long) As String
    'Hanyas = Hanyas + 1
    If IsNumeric(Mid(sor, 1, 1)) Then
            Sorszamoz = CStr(Hanyas) & ": " & SorszamNelkul(sor)
        Else
            Sorszamoz = CStr(Hanyas) & ": " & sor
    End If
End Function
Public Sub SorszamJavitasa(melyik As Long)
    'Melyik = Melyik - 1
    code.List(melyik) = Sorszamoz(code.List(melyik), melyik + 1)
    'kodablak.betolt Melyik, True
    'kodablak.ok_Click
End Sub
Public Function SorszamNelkul(sor As String) As String
    If InStr(1, sor, ": ") <> 0 Then
            SorszamNelkul = Mid(sor, InStr(1, sor, ": ") + 2)
        Else
            SorszamNelkul = sor
    End If
End Function
Public Sub Ujrasorszamoz()
Dim i As Long
    For i = 0 To code.ListCount - 1
        SorszamJavitasa (i)
    Next i
End Sub
Private Sub Frissit()
    'Aktuális utasítás
    akt_muv.Text = Utasitas(code.List(aktualis)) & " " & Parameter(code.List(aktualis), 1) & " " & Parameter(code.List(aktualis), 2)
    j_cim.Text = aktualis + 1
    kimenet.Text = kimeno
    
    'v_cim.Text = visszater
    'veremlista.Selected(0) = True
    
    'Regiszterek értékeinek kiírása
    reg(0).Text = X
    reg(1).Text = Y
    reg(2).Text = a
    reg(3).Text = b
    reg(4).Text = c
    reg(5).Text = d
    reg(6).Text = e
    reg(7).Text = f
    
    
    'Gombok kezelése
    Dim i As Long
    
    'Minden gomb alaphelyzetbe
    For i = 0 To menu.Count - 1
        menu(i).Enabled = True
        'menu(i).Top = 0
        menu(i).PaintPicture menu(i).Picture, 0, 0
    Next i
    
    'Gombok lokkolása futási státusz alapján
    code.Enabled = Not Fut
    
    For i = 0 To 5
        menu(i).Enabled = code.Enabled
    Next i
    
    menu(9).Enabled = code.Enabled
    menu(10).Enabled = code.Enabled
    
    'Lista állapot alapján
    If code.ListCount = 0 Or code.ListIndex = -1 Then
        menu(0).Enabled = False
        menu(2).Enabled = False
        menu(4).Enabled = False
        menu(5).Enabled = False
        menu(9).Enabled = False
        menu(10).Enabled = False
    End If
    
    If code.ListCount = 0 Then
        menu(6).Enabled = False
        menu(7).Enabled = False
        menu(8).Enabled = False
    End If
    
    If code.ListIndex = code.ListCount - 1 Then menu(10).Enabled = False
    
    If code.ListIndex = 0 Then menu(9).Enabled = False
    
    'For i = 0 To menu.Count - 1
        'If menu(i).Enabled Then
        '        menu(i).Cls
        '    Else
        '        Dim k As Integer
        '        For k = 1 To menu(i).Height Step 40
        '            menu(i).Line (0, k)-(menu(i).Width, k), vbWhite
        '        Next k
        'End If
    'Next i
    Me.Refresh
End Sub
Private Function Szamma(Mit) As Double
    If IsNumeric(Mit) Then
            Szamma = CDbl(Mit)
        Else
            Szamma = 0
    End If
End Function

Private Sub megnyitas_dlg()
On Error GoTo hiba
    pb.DialogTitle = szoveg(14)
    pb.Filter = szoveg(15) & "(*.txt)|*.txt|" & szoveg(16) & "(*.asm)|*.asm|" & szoveg(17) & "(*.*)|*.*"
    pb.ShowOpen
    
    megnyit pb.FileName, 0
hiba:
End Sub
Private Sub import_dlg()
On Error GoTo hiba
    pb.DialogTitle = szoveg(18)
    pb.Filter = szoveg(15) & "(*.txt)|*.txt|" & szoveg(16) & "(*.asm)|*.asm|" & szoveg(17) & "(*.*)|*.*"
    pb.ShowOpen
    
    megnyit pb.FileName, code.ListIndex + 1
hiba:
End Sub
Private Sub megnyit(FajlNev As String, Hova As Long)
Dim seged As String, i As Long
On Error GoTo hiba
    
    Open FajlNev For Input As 1
        If Hova = 0 Then code.Clear
        i = -1
        Do While Not EOF(1)
            Line Input #1, seged
            i = i + 1
            code.AddItem seged, Hova + i
            SorszamJavitasa (code.ListCount - 1)
        Loop
    Close 1
    If Hova <= 0 Then Me.Caption = nevem & " - " & FajlNev
    Frissit
    menu_Click (8)
    Ujrasorszamoz
Exit Sub

hiba:
    MsgBox Err.Description
    'Resume
    Close 1
End Sub
Private Sub mentes(FajlNev As String)
Dim i As Long
On Error GoTo hiba

    Open FajlNev For Output As 1
        For i = 0 To code.ListCount - 1
            Print #1, SorszamNelkul(code.List(i))
        Next i
    Close 1
    
    Me.Caption = nevem & " - " & FajlNev
Exit Sub
hiba:
    MsgBox Err.Description
    Close 1
End Sub
Private Sub mentes_dlg()
On Error GoTo hiba
    pb.DialogTitle = szoveg(19)
    pb.Filter = szoveg(15) & "(*.txt)|*.txt|" & szoveg(16) & "(*.asm)|*.asm"
    pb.ShowSave
    
    mentes (pb.FileName)
hiba:
End Sub

Private Sub nyelvvalaszto_Click()
    If nyelvvalaszto.ListIndex = 0 Then
            magyar_nyelv
        Else
            nyelv.nyelv (nyelvmappa.Path & "\" & nyelvvalaszto.Text & ".lng")
    End If
End Sub

Private Sub regpanel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub

Private Sub sebesseg_Click(Index As Integer)
    Select Case Index
        Case 0
            idozito.Interval = 1
        Case 1
            idozito.Interval = 1000
        Case 2
            idozito.Interval = 2000
        Case 3
            idozito.Interval = 5000
    End Select
End Sub


Private Property Get aktualis() As Long
    aktualis = aktualis_v '+ 1
End Property

Private Property Let aktualis(ByVal Ertek As Long)
On Error Resume Next
    aktualis_v = Ertek '- 1
    code.Selected(aktualis_v) = True
End Property

Private Sub vezerlo_tarto_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Form_MouseMove Button, Shift, X, Y
End Sub
Private Function VeremElso()
    VeremElso = veremlista.List(0)
End Function
Private Property Get Verem()
    Verem = VeremElso
    veremlista.RemoveItem (0)
End Property
Private Property Let Verem(ByVal Ertek)
    veremlista.AddItem Ertek, 0
End Property
Public Sub Loggol(szoveg As String, Optional NemKell As Boolean)
    If Not NemKell Then szoveg = aktualis + 1 & ": " & szoveg
    loglista.AddItem szoveg
End Sub
