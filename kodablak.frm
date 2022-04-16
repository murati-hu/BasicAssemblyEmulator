VERSION 5.00
Begin VB.Form kodablak 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Kód szerkesztõ"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5745
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "kodablak.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   5745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox p 
      Height          =   315
      Index           =   1
      ItemData        =   "kodablak.frx":0CCA
      Left            =   3000
      List            =   "kodablak.frx":0CE6
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.ComboBox p 
      Height          =   315
      Index           =   0
      ItemData        =   "kodablak.frx":0D02
      Left            =   1800
      List            =   "kodablak.frx":0D1E
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton megse 
      Caption         =   "Mégse"
      Height          =   375
      Left            =   4440
      TabIndex        =   5
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton ok 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4440
      TabIndex        =   4
      Top             =   240
      Width           =   1215
   End
   Begin VB.TextBox megj 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3975
   End
   Begin VB.ComboBox kod 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   360
      Width           =   1455
   End
   Begin VB.Label sugo 
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      Caption         =   "Megjegyzés:"
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   9
      Top             =   720
      Width           =   885
   End
   Begin VB.Line Line1 
      X1              =   4200
      X2              =   4200
      Y1              =   120
      Y2              =   2040
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      Caption         =   "2. Paraméter:"
      Height          =   195
      Index           =   2
      Left            =   3000
      TabIndex        =   8
      Top             =   120
      Width           =   945
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      Caption         =   "1. Paraméter:"
      Height          =   195
      Index           =   1
      Left            =   1800
      TabIndex        =   7
      Top             =   120
      Width           =   945
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      Caption         =   "Utasítás:"
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   630
   End
End
Attribute VB_Name = "kodablak"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private SorSzam As Long
Private letezik As Boolean
Public modosit As String, felvesz As String

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then megse_Click
End Sub

Private Sub Form_Load()
    kod.AddItem "REM"
    
    kod.AddItem "INP"
    kod.AddItem "LET"
    kod.AddItem "STR"
    
    kod.AddItem "ADD"
    kod.AddItem "INC"
    kod.AddItem "SUB"
    kod.AddItem "DEC"
    kod.AddItem "MLP"
    kod.AddItem "DIV"
    
    kod.AddItem "MOV"
    kod.AddItem "JMP"
    kod.AddItem "SIG"
    
    kod.AddItem "GTO"
    kod.AddItem "GSB"
    kod.AddItem "RET"
    
    kod.AddItem "OUT"
    kod.AddItem "END"
    
    kod.Text = kod.List(1)
    kod_Click
End Sub



Public Sub kod_Click()
Dim m As Byte
    m = 0
    Select Case kod.Text
        Case "REM"
            m = 2
            sugo.Caption = utasitasok(1)
        Case "INP"
            sugo.Caption = utasitasok(2)
        Case "LET"
            sugo.Caption = utasitasok(3)
        Case "STR"
            sugo.Caption = utasitasok(4)
            m = 1
        Case "ADD"
            sugo.Caption = utasitasok(5)
        Case "SUB"
            sugo.Caption = utasitasok(6)
        Case "INC"
            sugo.Caption = utasitasok(7)
            m = 1
        Case "DEC"
            sugo.Caption = utasitasok(8)
            m = 1
        Case "MLP"
            sugo.Caption = utasitasok(9)
        Case "DIV"
            sugo.Caption = utasitasok(10)
        Case "MOV"
            sugo.Caption = utasitasok(11)
        Case "JMP"
            sugo.Caption = utasitasok(12)
            m = 1
        Case "SIG"
            sugo.Caption = utasitasok(13)
            m = 1
        Case "GTO"
            sugo.Caption = utasitasok(14)
            m = 1
        Case "GSB"
            sugo.Caption = utasitasok(15)
            m = 1
        Case "RET"
            sugo.Caption = utasitasok(16)
            m = 1
        Case "OUT"
           sugo.Caption = utasitasok(17)
            m = 1
        Case "END"
            m = 2
            sugo.Caption = utasitasok(18)
    End Select
    
    p(0).Enabled = True
    p(1).Enabled = True
    If m = 1 Then p(1).Enabled = False
    If m = 2 Then
        p(0).Enabled = False
        p(1).Enabled = False
    End If
    
End Sub

Private Sub megse_Click()
    'Unload Me
    Me.Hide
    
End Sub

Public Sub ok_Click()
Dim sor As String
    sor = SorSzam + 1 & ": " & kod.Text & " " & p(0).Text & " " & p(1).Text
    
    If Len(Trim(megj.Text)) <> 0 Then
        sor = sor & " //" & megj.Text
    End If
    
    If letezik Then
            emulator.code.List(SorSzam) = sor
        Else
            Dim i As Long
            With emulator.code
                '.AddItem .List(.ListCount - 1) ', SorSzam
                'emulator.SorszamJavitasa (.ListCount - 1)
                'For i = .ListCount - 2 To SorSzam Step -1
                '    .List(i + 1) = .List(i)
                '    emulator.SorszamJavitasa (i + 1)
                'Next i
                '.List(SorSzam) = Sor
                ''emulator.SorszamJavitasa (SorSzam)
                
                .AddItem sor, SorSzam
                emulator.Ujrasorszamoz
                .Selected(SorSzam) = True
                
            End With
    End If
    megse_Click
End Sub

Public Sub Mutasd(Melyiket As Long, Van As Boolean)
'On Error Resume Next
    'Form_Load
    betolt Melyiket, Van
    kodablak.Show vbModal
End Sub
Public Sub betolt(Melyiket As Long, Van As Boolean)
On Error GoTo hiba
    letezik = Van
    If Van Then
            kodablak.kod.Text = emulator.Utasitas(emulator.code.List(Melyiket))
            kodablak.p(0).Text = emulator.Parameter(emulator.code.List(Melyiket), 1)
            kodablak.p(1).Text = emulator.Parameter(emulator.code.List(Melyiket), 2)
            If InStr(1, emulator.code.List(Melyiket), "//") <> 0 Then
                megj.Text = Mid(emulator.code.List(Melyiket), InStr(1, emulator.code.List(Melyiket), "//") + 2)
            End If
            'kodablak.megj.Text = emulator.Parameter(emulator.code.List(Melyiket), 3)
            'megj.Text = Mid(megj.Text, 3)
            ok.Caption = modosit
        Else
            ok.Caption = felvesz
    End If
    
    SorSzam = Melyiket
Exit Sub
hiba:
    MsgBox Err.Description, vbInformation, Err.Number
    Resume Next
End Sub

Private Sub p_Change(Index As Integer)
    p(Index).Text = UCase(p(Index).Text)
End Sub
