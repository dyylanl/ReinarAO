VERSION 5.00
Begin VB.Form frmCrearPersonaje 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "frmCrearPersonajedados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   0  'User
   ScaleWidth      =   12075.47
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox lstGenero 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonajedados.frx":324A
      Left            =   5160
      List            =   "frmCrearPersonajedados.frx":3254
      Style           =   2  'Dropdown List
      TabIndex        =   31
      Top             =   1950
      Width           =   2753
   End
   Begin VB.ComboBox lstRaza 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonajedados.frx":3267
      Left            =   5160
      List            =   "frmCrearPersonajedados.frx":327A
      Style           =   2  'Dropdown List
      TabIndex        =   30
      Top             =   1350
      Width           =   2753
   End
   Begin VB.ComboBox lstHogar 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmCrearPersonajedados.frx":32A7
      Left            =   5160
      List            =   "frmCrearPersonajedados.frx":32B4
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   2580
      Width           =   2753
   End
   Begin VB.TextBox txtNombre 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   465
      Left            =   480
      MaxLength       =   20
      TabIndex        =   0
      Top             =   600
      Width           =   6735
   End
   Begin VB.Label modCarisma 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6795
      TabIndex        =   39
      Top             =   7920
      Width           =   690
   End
   Begin VB.Label modInteligencia 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6720
      TabIndex        =   38
      Top             =   6960
      Width           =   735
   End
   Begin VB.Label modConstitucion 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6795
      TabIndex        =   37
      Top             =   6000
      Width           =   690
   End
   Begin VB.Label modAgilidad 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6795
      TabIndex        =   36
      Top             =   5040
      Width           =   735
   End
   Begin VB.Label modfuerza 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6795
      TabIndex        =   35
      Top             =   4080
      Width           =   855
   End
   Begin VB.Label lbSabiduria 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H00FFFF00&
      Height          =   255
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Visible         =   0   'False
      Width           =   270
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "+3"
      ForeColor       =   &H00FFFF80&
      Height          =   195
      Left            =   180
      TabIndex        =   33
      Top             =   0
      Visible         =   0   'False
      Width           =   195
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   21
      Left            =   11031
      TabIndex        =   29
      Top             =   7200
      Width           =   405
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   42
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":32D7
      MousePointer    =   99  'Custom
      Top             =   7290
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   43
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":3429
      MousePointer    =   99  'Custom
      Top             =   7305
      Width           =   195
   End
   Begin VB.Label puntosquedan 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "32"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   6240
      TabIndex        =   28
      Top             =   2955
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Puntos 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "10"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   10440
      TabIndex        =   27
      Top             =   360
      Width           =   270
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   3
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":357B
      MousePointer    =   99  'Custom
      Top             =   1185
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   5
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":36CD
      MousePointer    =   99  'Custom
      Top             =   1440
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   7
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":381F
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   9
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":3971
      MousePointer    =   99  'Custom
      Top             =   2070
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   11
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":3AC3
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   13
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":3C15
      MousePointer    =   99  'Custom
      Top             =   2700
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   15
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":3D67
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   17
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":3EB9
      MousePointer    =   99  'Custom
      Top             =   3270
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   19
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":400B
      MousePointer    =   99  'Custom
      Top             =   3615
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   21
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":415D
      MousePointer    =   99  'Custom
      Top             =   3945
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   23
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":42AF
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   25
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":4401
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   27
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":4553
      MousePointer    =   99  'Custom
      Top             =   4815
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   1
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":46A5
      MousePointer    =   99  'Custom
      Top             =   840
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   0
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":47F7
      MousePointer    =   99  'Custom
      Top             =   870
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   2
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":4949
      MousePointer    =   99  'Custom
      Top             =   1200
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   4
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":4A9B
      MousePointer    =   99  'Custom
      Top             =   1500
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   6
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":4BED
      MousePointer    =   99  'Custom
      Top             =   1800
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   8
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":4D3F
      MousePointer    =   99  'Custom
      Top             =   2085
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   10
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":4E91
      MousePointer    =   99  'Custom
      Top             =   2400
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   12
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":4FE3
      MousePointer    =   99  'Custom
      Top             =   2730
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   240
      Index           =   14
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":5135
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   16
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":5287
      MousePointer    =   99  'Custom
      Top             =   3360
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   18
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":53D9
      MousePointer    =   99  'Custom
      Top             =   3630
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   20
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":552B
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   22
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":567D
      MousePointer    =   99  'Custom
      Top             =   4200
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   24
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":57CF
      MousePointer    =   99  'Custom
      Top             =   4560
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   210
      Index           =   26
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":5921
      MousePointer    =   99  'Custom
      Top             =   4800
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   28
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":5A73
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   29
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":5BC5
      MousePointer    =   99  'Custom
      Top             =   5160
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   30
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":5D17
      MousePointer    =   99  'Custom
      Top             =   5490
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   31
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":5E69
      MousePointer    =   99  'Custom
      Top             =   5430
      Width           =   150
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   32
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":5FBB
      MousePointer    =   99  'Custom
      Top             =   5760
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   33
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":610D
      MousePointer    =   99  'Custom
      Top             =   5760
      Width           =   180
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   34
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":625F
      MousePointer    =   99  'Custom
      Top             =   6105
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   150
      Index           =   35
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":63B1
      MousePointer    =   99  'Custom
      Top             =   6090
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   225
      Index           =   36
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":6503
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   180
      Index           =   37
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":6655
      MousePointer    =   99  'Custom
      Top             =   6360
      Width           =   195
   End
   Begin VB.Image command1 
      Height          =   120
      Index           =   38
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":67A7
      MousePointer    =   99  'Custom
      Top             =   6720
      Width           =   135
   End
   Begin VB.Image command1 
      Height          =   165
      Index           =   39
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":68F9
      MousePointer    =   99  'Custom
      Top             =   6705
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   40
      Left            =   11400
      MouseIcon       =   "frmCrearPersonajedados.frx":6A4B
      MousePointer    =   99  'Custom
      Top             =   6990
      Width           =   165
   End
   Begin VB.Image command1 
      Height          =   135
      Index           =   41
      Left            =   10875
      MouseIcon       =   "frmCrearPersonajedados.frx":6B9D
      MousePointer    =   99  'Custom
      Top             =   6990
      Width           =   135
   End
   Begin VB.Image boton 
      Height          =   255
      Index           =   1
      Left            =   120
      MouseIcon       =   "frmCrearPersonajedados.frx":6CEF
      MousePointer    =   99  'Custom
      Top             =   8640
      Width           =   2685
   End
   Begin VB.Image boton 
      Appearance      =   0  'Flat
      Height          =   570
      Index           =   0
      Left            =   360
      MouseIcon       =   "frmCrearPersonajedados.frx":6E41
      MousePointer    =   99  'Custom
      Top             =   7680
      Width           =   4560
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   20
      Left            =   11031
      TabIndex        =   26
      Top             =   6900
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   19
      Left            =   11031
      TabIndex        =   25
      Top             =   6600
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   18
      Left            =   11031
      TabIndex        =   24
      Top             =   6285
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   17
      Left            =   11031
      TabIndex        =   23
      Top             =   5970
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   16
      Left            =   11031
      TabIndex        =   22
      Top             =   5685
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   15
      Left            =   11031
      TabIndex        =   21
      Top             =   5385
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   14
      Left            =   11031
      TabIndex        =   20
      Top             =   5070
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   13
      Left            =   11031
      TabIndex        =   19
      Top             =   4770
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   12
      Left            =   11031
      TabIndex        =   18
      Top             =   4470
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   11
      Left            =   11031
      TabIndex        =   17
      Top             =   4155
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   10
      Left            =   11031
      TabIndex        =   16
      Top             =   3840
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   9
      Left            =   11031
      TabIndex        =   15
      Top             =   3540
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   8
      Left            =   11031
      TabIndex        =   14
      Top             =   3225
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   7
      Left            =   11031
      TabIndex        =   13
      Top             =   2925
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   6
      Left            =   11031
      TabIndex        =   12
      Top             =   2610
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   5
      Left            =   11031
      TabIndex        =   11
      Top             =   2310
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   4
      Left            =   11031
      TabIndex        =   10
      Top             =   2010
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   3
      Left            =   11031
      TabIndex        =   9
      Top             =   1710
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   2
      Left            =   11031
      TabIndex        =   8
      Top             =   1395
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   0
      Left            =   11031
      TabIndex        =   7
      Top             =   780
      Width           =   398
   End
   Begin VB.Label Skill 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   300
      Index           =   1
      Left            =   11031
      TabIndex        =   6
      Top             =   1080
      Width           =   398
   End
   Begin VB.Label lbCarisma 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   450
      Left            =   6315
      TabIndex        =   5
      Top             =   7800
      Width           =   495
   End
   Begin VB.Label lbInteligencia 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   450
      Left            =   6315
      TabIndex        =   4
      Top             =   6840
      Width           =   495
   End
   Begin VB.Label lbConstitucion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   450
      Left            =   6315
      TabIndex        =   3
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label lbAgilidad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   450
      Left            =   6315
      TabIndex        =   2
      Top             =   4920
      Width           =   495
   End
   Begin VB.Label lbFuerza 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "18"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFF00&
      Height          =   450
      Left            =   6315
      TabIndex        =   1
      Top             =   3960
      Width           =   495
   End
End
Attribute VB_Name = "frmCrearPersonaje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public SkillPoints As Byte
Function CheckData() As Boolean
    If Len(UserName) >= 15 Then
        MsgBox "El nombre no puede tener mas de 15 caracteres."
        Exit Function
    End If

    If Len(UserName) <= 2 Then
        MsgBox "El nombre no puede tener menos de 2 caracteres."
        Exit Function
    End If
    If UCase$(UserName) = "NADA" Then
        MsgBox "El nombre es invalido."
        Exit Function
    End If

    If UserRaza = 0 Then
        MsgBox "Seleccione la raza del personaje."
        Exit Function
    End If

    If UserHogar = 0 Then
        MsgBox "Seleccione el hogar del personaje."
        Exit Function
    End If

    If UserSexo = -1 Then
        MsgBox "Seleccione el sexo del personaje."
        Exit Function
    End If

    If SkillPoints > 0 Then
        MsgBox "Asigne los skillpoints del personaje."
        Exit Function
    End If

    Dim i As Integer
    For i = 1 To NUMATRIBUTOS
        If UserAtributos(i) = 0 Then
            MsgBox "Los atributos del personaje son invalidos."
            Exit Function
        End If
    Next i

    CheckData = True

End Function
Private Sub boton_Click(Index As Integer)
    Dim i As Integer
    Dim k As Object
    Call SendData("TIRDAD")
    Call Sound.Sound_Play(SND_CLICK)

    Select Case Index
        Case 0
            LlegoConfirmacion = False
            Confirmacion = 0

            i = 1

            For Each k In Skill
                UserSkills(i) = k.Caption
                i = i + 1
            Next

            UserName = txtNombre.Text


            If Right$(UserName, 1) = " " Then
                UserName = Trim(UserName)
                MsgBox "Nombre invalido, se han removido los espacios al final del nombre"
            End If

            UserRaza = lstRaza.ListIndex + 1
            UserSexo = lstGenero.ListIndex
            UserHogar = lstHogar.ListIndex + 1

            UserAtributos(1) = 1
            UserAtributos(2) = 1
            UserAtributos(3) = 1
            UserAtributos(4) = 1
            UserAtributos(5) = 1

            If CheckData() Then
                frmMain.Socket1.HostName = IPdelServidor
                frmMain.Socket1.RemotePort = PuertoDelServidor

                Me.MousePointer = 11
                EstadoLogin = CrearNuevoPj

                'EncriptNeW txtNombre.Text, txtCorreo.Text

                If Not frmMain.Socket1.Connected Then
                    Call MsgBox("Error: Se ha perdido la conexion con el server.")
                    Unload Me
                Else
                    Call Login
                End If

                frmConnect.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\conectar.bmp")
            End If

        Case 1

            'frmConnect.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\conectar.bmp")

            frmMain.Socket1.Disconnect
            frmConnect.MousePointer = 1
            If Opciones.sMusica <> CONST_DESHABILITADA Then
                If Opciones.sMusica <> CONST_DESHABILITADA Then
                    Sound.NextMusic = MUS_VolverInicio
                    Sound.Fading = 350
                End If
            End If
            frmConnect.Show
            Unload Me
            
            

    End Select

End Sub
Private Sub Command1_Click(Index As Integer)
    Call Sound.Sound_Play(SND_CLICK)

    Dim indice
    If Index Mod 2 = 0 Then
        If SkillPoints > 0 Then
            indice = Index \ 2
            Skill(indice).Caption = Val(Skill(indice).Caption) + 1
            SkillPoints = SkillPoints - 1
        End If
    Else
        If SkillPoints < 10 Then

            indice = Index \ 2
            If Val(Skill(indice).Caption) > 0 Then
                Skill(indice).Caption = Val(Skill(indice).Caption) - 1
                SkillPoints = SkillPoints + 1
            End If
        End If
    End If

    Puntos.Caption = SkillPoints
End Sub
Private Sub Form_Load()

    SkillPoints = 10
    Puntos.Caption = SkillPoints
    Me.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\CrearPersonajeConDados.gif")
    Me.MousePointer = vbDefault

    Select Case (lstRaza.List(lstRaza.ListIndex))
        Case Is = "Humano"
            modfuerza.Caption = "+ 1"
            modConstitucion.Caption = "+ 2"
            modAgilidad.Caption = "+ 1"
            modInteligencia.Caption = ""
            modCarisma.Caption = ""
        Case Is = "Elfo"
            modfuerza.Caption = ""
            modConstitucion.Caption = "+ 1"
            modAgilidad.Caption = "+ 3"
            modInteligencia.Caption = "+ 1"
            modCarisma.Caption = "+ 2"
        Case Is = "Elfo Oscuro"
            modfuerza.Caption = "+ 1"
            modConstitucion.Caption = ""
            modAgilidad.Caption = "+ 1"
            modInteligencia.Caption = "+ 2"
            modCarisma.Caption = "- 3"
        Case Is = "Enano"
            modfuerza.Caption = "+ 3"
            modConstitucion.Caption = "+ 3"
            modAgilidad.Caption = "- 1"
            modInteligencia.Caption = "- 6"
            modCarisma.Caption = "- 3"
        Case Is = "Gnomo"
            modfuerza.Caption = "- 5"
            modAgilidad.Caption = "+ 4"
            modInteligencia.Caption = "+ 3"
            modCarisma.Caption = "+ 1"
    End Select

End Sub

Private Sub lstRaza_click()

    Select Case (lstRaza.List(lstRaza.ListIndex))
        Case Is = "Humano"
            modfuerza.Caption = "+ 1"
            modConstitucion.Caption = "+ 2"
            modAgilidad.Caption = "+ 1"
            modInteligencia.Caption = ""
            modCarisma.Caption = ""
        Case Is = "Elfo"
            modfuerza.Caption = ""
            modConstitucion.Caption = "+ 1"
            modAgilidad.Caption = "+ 3"
            modInteligencia.Caption = "+ 1"
            modCarisma.Caption = "+ 2"
        Case Is = "Elfo Oscuro"
            modfuerza.Caption = "+ 1"
            modConstitucion.Caption = ""
            modAgilidad.Caption = "+ 1"
            modInteligencia.Caption = "+ 2"
            modCarisma.Caption = "- 3"
        Case Is = "Enano"
            modfuerza.Caption = "+ 3"
            modConstitucion.Caption = "+ 3"
            modAgilidad.Caption = "- 1"
            modInteligencia.Caption = "- 6"
            modCarisma.Caption = "- 3"
        Case Is = "Gnomo"
            modfuerza.Caption = "- 5"
            modAgilidad.Caption = "+ 4"
            modInteligencia.Caption = "+ 3"
            modCarisma.Caption = "+ 1"
    End Select

End Sub

Private Sub txtNombre_Change()
    txtNombre.Text = LTrim(txtNombre.Text)
End Sub

Private Sub txtNombre_GotFocus()
    MsgBox "Sea cuidadoso al seleccionar el nombre de su personaje, Argentum es un juego de rol, un mundo magico y fantastico, si selecciona un nombre obsceno o con connotación politica los administradores borrarán su personaje y no habrá ninguna posibilidad de recuperarlo."
End Sub

