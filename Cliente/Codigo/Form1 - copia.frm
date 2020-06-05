VERSION 5.00
Begin VB.Form frmRECanjes 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Sistema de Recanjeo"
   ClientHeight    =   7695
   ClientLeft      =   420
   ClientTop       =   255
   ClientWidth     =   10425
   Icon            =   "Form1 - copia.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1 - copia.frx":57E2
   ScaleHeight     =   7695
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   8640
      TabIndex        =   10
      Top             =   6600
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Caption         =   "Recanjear"
      Height          =   375
      Left            =   5040
      MaskColor       =   &H00000000&
      TabIndex        =   9
      Top             =   6600
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      FillColor       =   &H00FFFFFF&
      ForeColor       =   &H00FFFFFF&
      Height          =   660
      Left            =   5040
      ScaleHeight     =   600
      ScaleWidth      =   600
      TabIndex        =   2
      Top             =   720
      Width           =   660
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   6300
      ItemData        =   "Form1 - copia.frx":21EBA
      Left            =   1680
      List            =   "Form1 - copia.frx":21EBC
      TabIndex        =   0
      Top             =   720
      Width           =   3135
   End
   Begin VB.Label lblPrecio 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   12
      Top             =   2280
      Width           =   2775
   End
   Begin VB.Label lblNombre 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   11
      Top             =   960
      Width           =   2775
   End
   Begin VB.Label lblPermisos 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   5040
      TabIndex        =   8
      Top             =   5040
      Width           =   3495
   End
   Begin VB.Label lblStat 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   7
      Top             =   3600
      Width           =   2775
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Clases e información"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   300
      Left            =   5880
      TabIndex        =   6
      Top             =   4680
      Width           =   1935
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Stats:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   300
      Left            =   5040
      TabIndex        =   5
      Top             =   3600
      Width           =   555
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Precio:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   300
      Left            =   4920
      TabIndex        =   4
      Top             =   2280
      Width           =   690
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Nombre:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   300
      Left            =   6720
      TabIndex        =   3
      Top             =   600
      Width           =   840
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Lista completa de items de canjeo de Lhirius AO para Recanjear. Selecciona el item a recanjear y clickea Recanjear."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   480
      Left            =   1680
      TabIndex        =   1
      Top             =   0
      Width           =   6855
   End
End
Attribute VB_Name = "frmRECanjes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Call Audio.PlayWave(SND_CLICK)
If List1.Text = "Tunica Faccionaria Gris (Altos)" Then Call SendData("/RECANJEO T1")
If List1.Text = "Tunica Faccionaria Roja (Bajos)" Then Call SendData("/RECANJEO T2")
If List1.Text = "Tunica Faccionaria Roja (Altos)" Then Call SendData("/RECANJEO T3")
If List1.Text = "Tunica Faccionaria Azul (Bajos)" Then Call SendData("/RECANJEO T4")
If List1.Text = "Tunica Faccionaria Azul (Altos)" Then Call SendData("/RECANJEO T5")
If List1.Text = "Daga Naturalista" Then Call SendData("/RECANJEO T6")
If List1.Text = "Baculo de los dioses" Then Call SendData("/RECANJEO T7")
If List1.Text = "Tunica Transparencias" Then Call SendData("/RECANJEO T8")
If List1.Text = "Chupines Grises" Then Call SendData("/RECANJEO T9")
If List1.Text = "Chupines Azules" Then Call SendData("/RECANJEO T10")
If List1.Text = "Chupines Rojos" Then Call SendData("/RECANJEO T11")
If List1.Text = "Tunica Amarrilla" Then Call SendData("/RECANJEO T12")
If List1.Text = "Tunica Celeste" Then Call SendData("/RECANJEO T13")
If List1.Text = "Tunica Blanca" Then Call SendData("/RECANJEO T14")
If List1.Text = "Tunica Viva" Then Call SendData("/RECANJEO T15")
If List1.Text = "Tunica Verde" Then Call SendData("/RECANJEO T16")
If List1.Text = "Armadura Soberana Blanca (Bajos)" Then Call SendData("/RECANJEO T17")
If List1.Text = "Armadura Soberana Blanca (Altos)" Then Call SendData("/RECANJEO T18")
If List1.Text = "Tunica de Rey (Bajos)" Then Call SendData("/RECANJEO T19")
If List1.Text = "Tunica de Rey (Altos)" Then Call SendData("/RECANJEO T20")
If List1.Text = "Daga +5" Then Call SendData("/RECANJEO T21")
If List1.Text = "Espada Clerical" Then Call SendData("/RECANJEO T22")
If List1.Text = "Espada Imperial" Then Call SendData("/RECANJEO T23")
If List1.Text = "Escudo Estrella" Then Call SendData("/RECANJEO T24")
If List1.Text = "Escudo Plaxus" Then Call SendData("/RECANJEO T25")
If List1.Text = "Sombrero de Mago Verde" Then Call SendData("/RECANJEO T26")
If List1.Text = "Sombrero de Mago Rojo" Then Call SendData("/RECANJEO T27")
If List1.Text = "Tiara de la Vida" Then Call SendData("/RECANJEO T28")
If List1.Text = "Coronita" Then Call SendData("/RECANJEO T29")
If List1.Text = "Corona Verde" Then Call SendData("/RECANJEO T30")
If List1.Text = "Coronita Dorada" Then Call SendData("/RECANJEO T31")
If List1.Text = "Corona de Rey" Then Call SendData("/RECANJEO T32")
If List1.Text = "Espada Fantasmal" Then Call SendData("/RECANJEO T33")
If List1.Text = "Arco Celestial" Then Call SendData("/RECANJEO T34")
If List1.Text = "Pendiente del Sacrificio" Then Call SendData("/RECANJEO T35")
End Sub

Private Sub Command2_Click()
Call Audio.PlayWave(SND_CLICK)
Unload Me
End Sub

Private Sub Form_Load()
List1.AddItem "Tunica Faccionaria Gris (Altos)"
List1.AddItem "Tunica Faccionaria Roja (Bajos)"
List1.AddItem "Tunica Faccionaria Roja (Altos)"
List1.AddItem "Tunica Faccionaria Azul (Bajos)"
List1.AddItem "Tunica Faccionaria Azul (Altos)"
List1.AddItem "Daga Naturalista"
List1.AddItem "Baculo de los dioses"
List1.AddItem "Tunica Transparencias"
List1.AddItem "Chupines Grises"
List1.AddItem "Chupines Azules"
List1.AddItem "Chupines Rojos"
List1.AddItem "Tunica Amarrilla"
List1.AddItem "Tunica Celeste"
List1.AddItem "Tunica Blanca"
List1.AddItem "Tunica Viva"
List1.AddItem "Tunica Verde"
List1.AddItem "Armadura Soberana Blanca (Bajos)"
List1.AddItem "Armadura Soberana Blanca (Altos)"
List1.AddItem "Tunica de Rey (Bajos)"
List1.AddItem "Tunica de Rey (Altos)"
List1.AddItem "Daga +5"
List1.AddItem "Espada Clerical"
List1.AddItem "Espada Fantasmal"
List1.AddItem "Espada Imperial"
List1.AddItem "Escudo Estrella"
List1.AddItem "Escudo Plaxus"
List1.AddItem "Sombrero de Mago Verde"
List1.AddItem "Sombrero de Mago Rojo"
List1.AddItem "Tiara de la Vida"
List1.AddItem "Coronita"
List1.AddItem "Corona Verde"
List1.AddItem "Coronita Dorada"
List1.AddItem "Corona de Rey"
List1.AddItem "Arco Celestial"
List1.AddItem "Pendiente del Sacrificio"
End Sub

Private Sub list1_Click()
Dim picture As String
If List1.Text = "Tunica Faccionaria Gris (Altos)" Then
    picture = "16119.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "25 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Clase/s: Todas las Clases - Razas: H/E/EO. - Informacion: Tunica especialmente diseñada por los dioses para caracterizar a los Neutrales."
End If

If List1.Text = "Tunica Faccionaria Roja (Bajos)" Then
    picture = "16126.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "25 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Clase/s: Todas las Clases - Razas: Gnomos y Enanos. -  Informacion: Tunica especialmente diseñada por Horda infernal para sus fieles Criminales."
End If

If List1.Text = "Tunica Faccionaria Roja (Altos)" Then
    picture = "16126.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "25 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Clase/s: Todas las Clases - Razas: H/E/EO. - Informacion: Tunica especialmente diseñada por Horda infernal para sus fieles Criminales."
End If

If List1.Text = "Tunica Faccionaria Azul (Bajos)" Then
    picture = "16129.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "25 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Clase/s: Todas las Clases - Razas: Gnomos y Enanos. - Informacion: Tunica especialmente diseñada por el Rey de Los Ciudadanos."
End If

If List1.Text = "Tunica Faccionaria Azul (Altos)" Then
    picture = "16129.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "25 Puntos de Canje"
    lblStat.Caption = "Min: 40 / Max: 45"
    lblPermisos.Caption = "Clase/s: Todas las Clases - Razas: H/E/EO. - Informacion: Tunica especialmente diseñada por el Rey de Los Ciudadanos."
End If

If List1.Text = "Daga Naturalista" Then
    picture = "1032.png"
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 8 / Max: 10"
    lblPermisos.Caption = "Clase/s: Bardo y Druida - Razas: Todas las razas. - Informacion: Arma especialmente diseñada por los dioses del mundo de Lhirius para Bardos y Druidas con gran ataque."
End If

If List1.Text = "Baculo de los dioses" Then
    picture = "16281.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 5 / Max: 7"
    lblPermisos.Caption = "Clase/s: Mago, Nigromante. - Razas: Todas las razas. - Informacion: Tiene un buen ataque magico."
End If

If List1.Text = "Tunica Transparencias" Then
    picture = "16136.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "50 Puntos de Canje"
    lblStat.Caption = "Min: 50 / Max: 55"
    lblPermisos.Caption = "Clase/s: Todas las Clases - Razas: Todas las razas."
End If

If List1.Text = "Chupines Grises" Then
    picture = "16184.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 45 / Max: 50"
    lblPermisos.Caption = "Clase/s: Todas las Clases - Razas: Todas las razas."
End If

If List1.Text = "Chupines Azules" Then
    picture = "16182.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 45 / Max: 50"
    lblPermisos.Caption = "Clase/s: Todas las Clases - Razas: Todas las razas."
End If

If List1.Text = "Chupines Rojos" Then
    picture = "16180.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 45 / Max: 50"
    lblPermisos.Caption = "Clase/s: Todas las Clases - Razas: Todas las razas."
End If

If List1.Text = "Tunica Amarrilla" Then
    picture = "16214.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "50 Puntos de Canje"
    lblStat.Caption = "Min: 50 / Max: 55"
    lblPermisos.Caption = "Clase/s: Todas las Clases - Razas: Todas las razas."
End If

If List1.Text = "Tunica Celeste" Then
    picture = "16216.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "50 Puntos de Canje"
    lblStat.Caption = "Min: 50 / Max: 55"
    lblPermisos.Caption = "Todas las Clases y Razas"
End If

If List1.Text = "Tunica Blanca" Then
    picture = "16231.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "50 Puntos de Canje"
    lblStat.Caption = "Min: 50 / Max: 55"
    lblPermisos.Caption = "Clase/s: Todas las Clases - Razas: Todas las razas."
End If

If List1.Text = "Tunica Viva" Then
    picture = "16233.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "50 Puntos de Canje"
    lblStat.Caption = "Min: 50 / Max: 55"
    lblPermisos.Caption = "Clase/s: Todas las Clases - Razas: Todas las razas."
End If

If List1.Text = "Tunica Verde" Then
    picture = "16235.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "50 Puntos de Canje"
    lblStat.Caption = "Min: 50 / Max: 55"
    lblPermisos.Caption = "Clase/s: Todas las Clases - Razas: Todas las razas"
End If

If List1.Text = "Armadura Soberana Blanca (Bajos)" Then
    picture = "16123.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 60 / Max: 65"
    lblPermisos.Caption = "Clase/s: Paladin, Guerrero, Cazador, Arquero - Razas: Gnomo y Enanos"
End If

If List1.Text = "Armadura Soberana Blanca (Altos)" Then
    picture = "16123.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 60 / Max: 65"
    lblPermisos.Caption = "Clase/s: Paladin, Guerrero, Cazador, Arquero - Razas: H/E/EO"
End If

If List1.Text = "Tunica de Rey (Bajos)" Then
    picture = "16089.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "4 Puntos de Canje"
    lblStat.Caption = "Min: 30 / Max: 35"
    lblPermisos.Caption = "Clase/s: Todas las Clases - Razas: Gnomos y Enanos - Informacion: Tunica del Difunto Rey del mundo Lhirius que fue heredada para los mas nobles."
End If

If List1.Text = "Tunica de Rey (Altos)" Then
    picture = "16089.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "4 Puntos de Canje"
    lblStat.Caption = "Min: 30 / Max: 35"
    lblPermisos.Caption = "Clase/s: Todas las Clases - Razas: H/E/EO - Informacion: Tunica del Difunto Rey del mundo Lhirius que fue heredada para los mas nobles."
End If

If List1.Text = "Daga +5" Then
    picture = "16150.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "50 Puntos de Canje"
    lblStat.Caption = "Min: 8 / Max: 12"
    lblPermisos.Caption = "Clase/s Asesino - Razas: Todas las razas. - Informacion: La daga mas poderosa que existe para Asesinos"
End If

If List1.Text = "Espada Clerical" Then
    picture = "16200.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "50 Puntos de Canje"
    lblStat.Caption = "Min: 18 / Max: 22"
    lblPermisos.Caption = "Clase/s: Clerigos - Razas: Todas las razas. -  Informacion: Espada del mejor Clerigo del mundo Lhirius que fue heredada para los mejores y mas poderosos Clerigos."
End If

If List1.Text = "Espada Fantasmal" Then
    picture = "9630.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "50 Puntos de Canje"
    lblStat.Caption = "Min: 20 / Max: 25"
    lblPermisos.Caption = "Clase/s: Guerreros - Razas: Todas las razas. - Informacion: Espada que pertenecio a un Noble Guerrero que heredo esta espada en manos de los mas fuertes."
End If

If List1.Text = "Espada Imperial" Then
    picture = "16083.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "50 Puntos de Canje"
    lblStat.Caption = "Min: 20 / Max: 22"
    lblPermisos.Caption = "Clase/s: Paladines -  Razas: Todas las razas. - Informacion: Espada diseñada por los imperios romanos especialmente para la utilizacion de paladines."
End If

If List1.Text = "Escudo Estrella" Then
    picture = "16152.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 20 / Max: 25"
    lblPermisos.Caption = "Clase/s: Guerrero, Paladin, Clerigo - Razas: Todas las razas - Informacion: Una imponente defenza."
End If

If List1.Text = "Escudo Plaxus" Then
    picture = "16154.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 15 / Max: 20"
    lblPermisos.Caption = "Clase/s: Asesino, Bardo y Druida - Razas: Todas las razas. - Informacion: Gran defenza en cuerpo a cuerpo para las clases semi-magicas."
End If

If List1.Text = "Sombrero de Mago Verde" Then
    picture = "16253.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 15 / Max: 17"
    lblPermisos.Caption = "Clase/s: Mago y Nigromante, Razas: Todas las razas. - Informacion: Una gran defenza magica."
End If

If List1.Text = "Sombrero de Mago Rojo" Then
    picture = "16237.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 15 / Max: 17"
    lblPermisos.Caption = "Clase/s: Mago y Nigromante, Razas: Todas las razas. - Informacion: Una gran defenza magica."
End If

If List1.Text = "Tiara de la Vida" Then
    picture = "16226.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "5 Puntos de Canje"
    lblStat.Caption = "Min: 25 / Max: 30"
    lblPermisos.Caption = "Clase/s : Todas las clases - Razas: Todas las razas"
End If

If List1.Text = "Coronita" Then
    picture = "16190.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "20 Puntos de Canje"
    lblStat.Caption = "Min: 13 / Max: 15"
    lblPermisos.Caption = "Clase/s: Todas las Clases. - Razas: Todas las Razas. - Informacion: Efectivo en defenza magica."
End If

If List1.Text = "Corona Verde" Then
    picture = "16245.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 45 / Max: 50"
    lblPermisos.Caption = "Clase/s: Todas las Clases. - Razas: Todas las Razas. - Informacion: Efectivo en combate contra paladines o guerreros"
End If

If List1.Text = "Coronita Dorada" Then
    picture = "16255.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 45 / Max: 50"
    lblPermisos.Caption = "Clase/s: Todas las Clases. - Razas: Todas las Razas. - Informacion: Efectivo en combate contra paladines o guerreros"
End If

If List1.Text = "Corona de Rey" Then
    picture = "16192.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "30 Puntos de Canje"
    lblStat.Caption = "Min: 45 / Max: 50"
    lblPermisos.Caption = "Clase/s: Todas las Clases. - Razas: Todas las Razas. - Informacion: Efectivo en combate contra paladines o guerreros"
End If

If List1.Text = "Arco Celestial" Then
    picture = "16198.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "45 Puntos de Canje"
    lblStat.Caption = "Min: 18 / Max: 20"
    lblPermisos.Caption = "Clase/s: Cazador, Arquero. - Razas: Todas las razas. - Informacion: El arco mas poderoso que existe en el mundo Lhirius para los mas poderosos habiles con arcos y proyectiles."
    End If
    
If List1.Text = "Pendiente del Sacrificio" Then
    picture = "16633.png"
    lblNombre.Caption = List1.Text
    lblPrecio.Caption = "10 Puntos de Canje"
    lblStat.Caption = "Min: N/A / Max: N/A"
    lblPermisos.Caption = "Clase/s: Todas. - Razas: Todas las razas. - Informacion: El pendiente que permite que no se caigan los items cuando un usuario muere al que posee este objeto. (1 USO)"
End If

If Extract_File(Graphics, App.Path & "\GRAFICOS\", picture, App.Path & "\GRAFICOS\") Then
        Call PngPictureLoad(App.Path & "\GRAFICOS\" & picture, Picture1, False)
        Call Kill(App.Path & "\Graficos\*.png")
    End If
End Sub
