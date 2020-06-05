VERSION 5.00
Begin VB.Form frmGuildBrief 
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7635
   ControlBox      =   0   'False
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Desc 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   12
      Top             =   5640
      Width           =   6495
   End
   Begin VB.Image aliado 
      Height          =   495
      Left            =   2760
      MouseIcon       =   "frmGuildBrief.frx":0000
      MousePointer    =   99  'Custom
      Top             =   6960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image command3 
      Height          =   495
      Left            =   5160
      MouseIcon       =   "frmGuildBrief.frx":030A
      MousePointer    =   99  'Custom
      Top             =   6960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image guerra 
      Height          =   495
      Left            =   360
      MouseIcon       =   "frmGuildBrief.frx":0614
      MousePointer    =   99  'Custom
      Top             =   6960
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image command1 
      Height          =   375
      Left            =   6840
      MouseIcon       =   "frmGuildBrief.frx":091E
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   855
   End
   Begin VB.Image command2 
      Height          =   495
      Left            =   2040
      MouseIcon       =   "frmGuildBrief.frx":0C28
      MousePointer    =   99  'Custom
      Top             =   6960
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.Label nombre 
      BackStyle       =   0  'Transparent
      Caption         =   "Nombre:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1440
      TabIndex        =   11
      Top             =   840
      Width           =   5535
   End
   Begin VB.Label fundador 
      BackStyle       =   0  'Transparent
      Caption         =   "Fundador:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   10
      Top             =   1080
      Width           =   5415
   End
   Begin VB.Label creacion 
      BackStyle       =   0  'Transparent
      Caption         =   "Fecha de creacion:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   9
      Top             =   1320
      Width           =   4335
   End
   Begin VB.Label lider 
      BackStyle       =   0  'Transparent
      Caption         =   "Lider:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   8
      Top             =   1680
      Width           =   5745
   End
   Begin VB.Label web 
      BackStyle       =   0  'Transparent
      Caption         =   "Web site:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   7
      Top             =   1920
      Width           =   5265
   End
   Begin VB.Label Miembros 
      BackStyle       =   0  'Transparent
      Caption         =   "Miembros:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1680
      TabIndex        =   6
      Top             =   2160
      Width           =   5355
   End
   Begin VB.Label eleccion 
      BackStyle       =   0  'Transparent
      Caption         =   "Dias para proxima eleccion de lider:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4080
      TabIndex        =   5
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Enemigos 
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes Enemigos:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2400
      TabIndex        =   4
      Top             =   2475
      Width           =   4875
   End
   Begin VB.Label Aliados 
      BackStyle       =   0  'Transparent
      Caption         =   "Clanes Aliados:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   3
      Top             =   2760
      Width           =   4950
   End
   Begin VB.Label Codex 
      BackColor       =   &H80000012&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   0
      Left            =   600
      TabIndex        =   2
      Top             =   3960
      Width           =   6495
   End
   Begin VB.Label Codex 
      BackColor       =   &H80000012&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   1
      Left            =   600
      TabIndex        =   1
      Top             =   4200
      Width           =   6495
   End
   Begin VB.Label Codex 
      BackColor       =   &H80000012&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Index           =   2
      Left            =   600
      TabIndex        =   0
      Top             =   4440
      Width           =   6495
   End
End
Attribute VB_Name = "frmGuildBrief"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FénixAO en DX8 by ·Parra, Thusing y DarkTester

Public EsLeader As Boolean
Public Sub ParseGuildInfo(ByVal buffer As String)
    Dim BandoClan As Byte, BandoMio As Byte

    BandoClan = Val(ReadField(8, buffer, Asc("¬")))
    BandoMio = Val(ReadField(11, buffer, Asc("¬")))

    If Not EsLeader Then
        Me.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\DetallesDeClan.jpg")
        guerra.Visible = False
        aliado.Visible = False
        Command3.Visible = False
        Command2.Visible = (BandoMio = BandoClan)
    Else
        Me.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\DetallesDeClanGuildMaster.gif")
        aliado.Visible = True
        guerra.Visible = True
        Command3.Visible = True
        Command2.Visible = False
    End If

    Select Case BandoClan
        Case 1
            Nombre.ForeColor = &HFF0000
        Case 2
            Nombre.ForeColor = &HFF&
        Case Else
            Nombre.ForeColor = &HE0E0E0
    End Select

    Nombre.Caption = ReadField(1, buffer, Asc("¬"))
    fundador.Caption = ReadField(2, buffer, Asc("¬"))
    creacion.Caption = ReadField(3, buffer, Asc("¬"))
    lider.Caption = ReadField(4, buffer, Asc("¬"))
    web.Caption = ReadField(5, buffer, Asc("¬"))
    Miembros.Caption = ReadField(6, buffer, Asc("¬"))
    Eleccion.Caption = ReadField(7, buffer, Asc("¬"))
    Enemigos.Caption = ReadField(9, buffer, Asc("¬"))
    Aliados.Caption = ReadField(10, buffer, Asc("¬"))

    Dim T%, k%
    k% = Val(ReadField(12, buffer, Asc("¬")))

    For T% = 1 To 3
        Codex(T% - 1).Caption = ReadField(12 + T%, buffer, Asc("¬"))
    Next T%
    Dim des$


    des$ = ReadField(12 + T%, buffer, Asc("¬"))

    desc = Replace(des$, "º", vbCrLf)

    Me.Show vbModeless, frmMain

End Sub

Private Sub aliado_Click()
    Call SendData("DECALIAD" & Right$(Nombre, Len(Nombre.Caption)))
    Unload Me
End Sub
Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Command2_Click()

    Call frmGuildSol.RecieveSolicitud(Right$(Nombre, Len(Nombre.Caption)))
    Call frmGuildSol.Show(vbModeless, frmGuildBrief)


End Sub

Private Sub Command3_Click()
    frmCommet.Nombre = Right$(Nombre.Caption, Len(Nombre.Caption))
    Call frmCommet.Show(vbModeless, frmGuildBrief)

End Sub

Private Sub guerra_Click()
    Call SendData("DECGUERR" & Right$(Nombre.Caption, Len(Nombre.Caption)))
    Unload Me
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bmoving = False And Button = vbLeftButton Then
        Dx3 = X
        dy = Y
        bmoving = True
    End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bmoving And ((X <> Dx3) Or (Y <> dy)) Then Move Left + (X - Dx3), Top + (Y - dy)

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then bmoving = False

End Sub
