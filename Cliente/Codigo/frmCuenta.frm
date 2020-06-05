VERSION 5.00
Begin VB.Form frmCuent 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   9000
   ScaleWidth      =   11985
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label CP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear Nuevo Personaje"
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   7
      Left            =   8760
      TabIndex        =   24
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label CP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear Nuevo Personaje"
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   6
      Left            =   6600
      TabIndex        =   23
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label CP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear Nuevo Personaje"
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   5
      Left            =   4440
      TabIndex        =   22
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label CP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear Nuevo Personaje"
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Index           =   4
      Left            =   2400
      TabIndex        =   21
      Top             =   5640
      Width           =   975
   End
   Begin VB.Label CP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear Nuevo Personaje"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Index           =   3
      Left            =   8760
      TabIndex        =   20
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label CP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear Nuevo Personaje"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Index           =   2
      Left            =   6600
      TabIndex        =   19
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label CP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear Nuevo Personaje"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Index           =   1
      Left            =   4440
      TabIndex        =   18
      Top             =   2880
      Width           =   975
   End
   Begin VB.Label personaje 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   2280
      TabIndex        =   16
      Top             =   1800
      Width           =   7455
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   8280
      Top             =   8160
      Width           =   3375
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   4200
      Top             =   8040
      Width           =   3615
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Top             =   8040
      Width           =   3615
   End
   Begin VB.Label lblNivel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   7
      Left            =   8280
      TabIndex        =   15
      Top             =   7530
      Width           =   1815
   End
   Begin VB.Label lblNivel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   6
      Left            =   6200
      TabIndex        =   14
      Top             =   7530
      Width           =   1815
   End
   Begin VB.Label lblNivel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   5
      Left            =   4050
      TabIndex        =   13
      Top             =   7530
      Width           =   1815
   End
   Begin VB.Label lblNivel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   4
      Left            =   1950
      TabIndex        =   12
      Top             =   7530
      Width           =   1815
   End
   Begin VB.Label lblNivel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   3
      Left            =   8280
      TabIndex        =   11
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblNivel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   2
      Left            =   6200
      TabIndex        =   10
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblNivel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   1
      Left            =   4050
      TabIndex        =   9
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblNivel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   255
      Index           =   0
      Left            =   1950
      TabIndex        =   8
      Top             =   4800
      Width           =   1815
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   7
      Left            =   8280
      TabIndex        =   7
      Top             =   7210
      Width           =   1815
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   6
      Left            =   6200
      TabIndex        =   6
      Top             =   7210
      Width           =   1815
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   5
      Left            =   4050
      TabIndex        =   5
      Top             =   7210
      Width           =   1815
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   4
      Left            =   1950
      TabIndex        =   4
      Top             =   7210
      Width           =   1815
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   3
      Left            =   8280
      TabIndex        =   3
      Top             =   4530
      Width           =   1815
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   6200
      TabIndex        =   2
      Top             =   4530
      Width           =   1815
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   1
      Left            =   4050
      TabIndex        =   1
      Top             =   4530
      Width           =   1815
   End
   Begin VB.Label lblNombre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Nada"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   1950
      TabIndex        =   0
      Top             =   4530
      Width           =   1815
   End
   Begin VB.Image ImgClase 
      Height          =   1590
      Index           =   7
      Left            =   8560
      Top             =   5540
      Width           =   1275
   End
   Begin VB.Image ImgClase 
      Height          =   1590
      Index           =   6
      Left            =   6440
      Top             =   5540
      Width           =   1275
   End
   Begin VB.Image ImgClase 
      Height          =   1590
      Index           =   5
      Left            =   4290
      Top             =   5540
      Width           =   1275
   End
   Begin VB.Image ImgClase 
      Height          =   1590
      Index           =   4
      Left            =   2220
      Top             =   5540
      Width           =   1275
   End
   Begin VB.Image ImgClase 
      Height          =   1590
      Index           =   3
      Left            =   8560
      Top             =   2835
      Width           =   1275
   End
   Begin VB.Image ImgClase 
      Height          =   1590
      Index           =   2
      Left            =   6440
      Top             =   2835
      Width           =   1275
   End
   Begin VB.Image ImgClase 
      Height          =   1590
      Index           =   1
      Left            =   4290
      Top             =   2840
      Width           =   1275
   End
   Begin VB.Image ImgClase 
      Height          =   1590
      Index           =   0
      Left            =   2220
      Top             =   2840
      Width           =   1275
   End
   Begin VB.Label CP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Crear Nuevo Personaje"
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Index           =   0
      Left            =   2400
      TabIndex        =   17
      Top             =   2880
      Width           =   975
   End
End
Attribute VB_Name = "frmCuent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CP_Click(Index As Integer)
    Call Sound.Sound_Play(SND_CLICK)

    If lblNombre(7).Caption <> "Nada" Then
        MsgBox "Tu cuenta ha llegado al máximo de personajes."
        Exit Sub
    End If

    EstadoLogin = Dados
    frmCrearPersonaje.Show ' vbModal
    Me.MousePointer = 11
    Unload Me

End Sub

Private Sub Form_Load()
    PJClickeado = "Nada"
    personaje.Caption = "Seleccione un personaje."
    Me.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\PanelPj.jpg")
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    For i = 0 To 7
        lblNombre(i).ForeColor = &HFF&
        lblNivel(i).ForeColor = &H40C0&
    Next i

End Sub

Private Sub Image1_Click()
    frmMain.Socket1.Disconnect
    frmConnect.MousePointer = 1
    frmConnect.Show
    Unload Me
End Sub

Private Sub Image2_Click()

    Call Sound.Sound_Play(SND_CLICK)

    If lblNombre(7).Caption <> "Nada" Then
        MsgBox "Tu cuenta ha llegado al máximo de personajes."
        Exit Sub
    End If

    EstadoLogin = Dados
    frmCrearPersonaje.Show ' vbModal
    Me.MousePointer = 11
    Unload Me
End Sub

Private Sub Image3_Click()
    If PJClickeado = "Nada" Then
        MsgBox "Seleccione un pj"
        Exit Sub
    End If
    Call Sound.Sound_Play(SND_CLICK)
    UserName = PJClickeado
    personaje.Caption = PJClickeado
    EstadoLogin = Normal
    Call Login
    Unload Me
End Sub

Private Sub ImgClase_Click(Index As Integer)
    PJClickeado = lblNombre(Index).Caption
    If PJClickeado = "Nada" Then
        personaje.Caption = "Seleccione un personaje."
    Else
        personaje.Caption = PJClickeado
    End If
End Sub

Private Sub ImgClase_DblClick(Index As Integer)
    If PJClickeado = "Nada" Then Exit Sub

    Call Sound.Sound_Play(SND_CLICK)
    UserName = PJClickeado
    personaje.Caption = PJClickeado
    EstadoLogin = Normal
    Call Login
    Unload Me
End Sub

Private Sub ImgClase_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblNombre(Index).ForeColor = &HC0C0&
    lblNivel(Index).ForeColor = &HC0C0&
End Sub

Private Sub lblNivel_Click(Index As Integer)
    PJClickeado = lblNombre(Index).Caption

    If PJClickeado = "Nada" Then
        personaje.Caption = "Seleccione un personaje."
    Else
        personaje.Caption = PJClickeado
    End If
End Sub

Private Sub lblNivel_DblClick(Index As Integer)
    If PJClickeado = "Nada" Then Exit Sub

    Call Sound.Sound_Play(SND_CLICK)
    UserName = PJClickeado
    personaje.Caption = PJClickeado
    EstadoLogin = Normal
    Call Login
    Unload Me

End Sub

Private Sub lblNivel_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblNombre(Index).ForeColor = &HC0C0&
    lblNivel(Index).ForeColor = &HC0C0&
End Sub

Private Sub lblNombre_Click(Index As Integer)
    PJClickeado = lblNombre(Index).Caption
    If PJClickeado = "Nada" Then
        personaje.Caption = "Seleccione un personaje."
    Else
        personaje.Caption = PJClickeado
    End If
End Sub

Private Sub lblNombre_DblClick(Index As Integer)
    If PJClickeado = "Nada" Then Exit Sub

    Call Sound.Sound_Play(SND_CLICK)
    UserName = PJClickeado
    personaje.Caption = PJClickeado
    EstadoLogin = Normal
    Call Login
    Unload Me

End Sub

Private Sub lblNombre_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    lblNombre(Index).ForeColor = &HC0C0&
    lblNivel(Index).ForeColor = &HC0C0&
End Sub
