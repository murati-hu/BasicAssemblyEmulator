VERSION 5.00
Begin VB.Form nevjegy 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Form1"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4470
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   4470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox logo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   900
      Left            =   0
      Picture         =   "nevjegy.frx":0000
      ScaleHeight     =   900
      ScaleWidth      =   4500
      TabIndex        =   3
      Top             =   0
      Width           =   4500
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   4440
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Label url 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "muratiakos@hotmail.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   195
      Index           =   1
      Left            =   1320
      TabIndex        =   10
      Top             =   2640
      Width           =   2115
   End
   Begin VB.Label url 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.ase.ini.hu"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   195
      Index           =   0
      Left            =   1320
      TabIndex        =   9
      Top             =   2280
      Width           =   1875
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   2640
      Width           =   585
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Web:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   465
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright Muráti Ákos 2004"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   195
      Index           =   6
      Left            =   960
      TabIndex        =   6
      Top             =   900
      Width           =   2355
   End
   Begin VB.Line Line1 
      BorderWidth     =   5
      X1              =   0
      X2              =   4440
      Y1              =   1150
      Y2              =   1150
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Köszönet:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   3
      Left            =   120
      TabIndex        =   5
      Top             =   3000
      Width           =   855
   End
   Begin VB.Label cimke 
      BackStyle       =   0  'Transparent
      Caption         =   "Eredeti magyar változat"
      ForeColor       =   &H0000C000&
      Height          =   435
      Index           =   0
      Left            =   1320
      TabIndex        =   4
      Top             =   3360
      Width           =   3105
   End
   Begin VB.Label koszonet 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Rozgonyi-Borus Ferenc"
      ForeColor       =   &H0000C000&
      Height          =   195
      Left            =   1320
      TabIndex        =   2
      Top             =   3000
      Width           =   1650
   End
   Begin VB.Label cimke 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Fordítás:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   238
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   195
      Index           =   2
      Left            =   120
      TabIndex        =   1
      Top             =   3360
      Width           =   765
   End
   Begin VB.Label cimke 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "8 regiszteres egyszerûsített assembly fejlesztõi környezet."
      ForeColor       =   &H0000C000&
      Height          =   975
      Index           =   1
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   4215
   End
End
Attribute VB_Name = "nevjegy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cimke_Click(Index As Integer)
    Form_Click
End Sub

Private Sub cimke_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Nincsalahuzva
End Sub

Private Sub Form_Click()
    Me.Hide
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Nincsalahuzva
End Sub

Private Sub logo_Click()
    Form_Click
End Sub

Private Sub url_Click(Index As Integer)
On Error Resume Next
    If Index = 0 Then
        Shell "C:\Program Files\Internet Explorer\iexplore.exe http://www.ase.ini.hu", vbNormalFocus
    Else
        Shell "C:\Program Files\Internet Explorer\iexplore.exe mailto:muratiakos@hotmail.com", vbNormalFocus
    End If
End Sub

Private Sub url_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    url(Index).FontUnderline = True
End Sub
Public Sub Nincsalahuzva()
    url(0).FontUnderline = False
    url(1).FontUnderline = False
End Sub
