VERSION 5.00
Begin VB.Form FrmElegirCamino 
   BackColor       =   &H80000002&
   BorderStyle     =   0  'None
   ClientHeight    =   7425
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7185
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   7185
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "Más información"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4800
      MouseIcon       =   "FrmElegirCamino.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   9
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "Más información"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3000
      MouseIcon       =   "FrmElegirCamino.frx":030A
      MousePointer    =   99  'Custom
      TabIndex        =   8
      Top             =   6360
      Width           =   1215
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   "Más información"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1080
      MouseIcon       =   "FrmElegirCamino.frx":0614
      MousePointer    =   99  'Custom
      TabIndex        =   7
      Top             =   3600
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   0
      MouseIcon       =   "FrmElegirCamino.frx":091E
      MousePointer    =   99  'Custom
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Image Fidelidad 
      Height          =   495
      Index           =   2
      Left            =   4680
      MouseIcon       =   "FrmElegirCamino.frx":0C28
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Image Fidelidad 
      Height          =   495
      Index           =   1
      Left            =   840
      MouseIcon       =   "FrmElegirCamino.frx":0F32
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   1575
   End
   Begin VB.Image command3 
      Height          =   495
      Left            =   2640
      MouseIcon       =   "FrmElegirCamino.frx":123C
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   1695
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      BackStyle       =   0  'Transparent
      Caption         =   "Mantenerse neutral"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   930
      TabIndex        =   6
      Top             =   4610
      Width           =   5415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmElegirCamino.frx":1546
      ForeColor       =   &H0000FFFF&
      Height          =   1455
      Left            =   900
      TabIndex        =   5
      Top             =   4950
      Width           =   5445
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmElegirCamino.frx":16FC
      ForeColor       =   &H000000C0&
      Height          =   1335
      Left            =   3960
      TabIndex        =   4
      Top             =   2160
      Width           =   2805
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmElegirCamino.frx":1802
      ForeColor       =   &H000000C0&
      Height          =   1455
      Left            =   240
      TabIndex        =   3
      Top             =   2160
      Width           =   2880
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"FrmElegirCamino.frx":18FC
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   960
      TabIndex        =   2
      Top             =   480
      Width           =   5415
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fiel a la Horda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Fiel a la Alianza"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
End
Attribute VB_Name = "FrmElegirCamino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command3_Click()
    Call SendData("SF0")
    Unload Me
End Sub
Private Sub Fidelidad_Click(Index As Integer)

    Unload frmfidelidad
    Fide = Index
    frmfidelidad.Show

End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Suclases3op.gif")
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
Private Sub Image1_Click()

    Unload Me

End Sub
Private Sub Label10_Click()
    Ayuda = 1
    SubAyuda = 2
    FrmAyuda.Show
End Sub

Private Sub Label8_Click()
    Ayuda = 1
    SubAyuda = 1
    FrmAyuda.Show
End Sub

Private Sub Label9_Click()
    Ayuda = 1
    SubAyuda = 3
    FrmAyuda.Show
End Sub
