VERSION 5.00
Begin VB.Form frmComandosGM 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Multi-Panel del Game Master"
   ClientHeight    =   7905
   ClientLeft      =   6660
   ClientTop       =   810
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   8700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame13 
      BackColor       =   &H00FFFF80&
      Caption         =   "Torneos"
      ForeColor       =   &H00FF0000&
      Height          =   1575
      Left            =   6960
      TabIndex        =   102
      Top             =   6480
      Width           =   1695
      Begin VB.CommandButton Command88 
         Caption         =   "Torneo 1vs1"
         Height          =   375
         Left            =   240
         TabIndex        =   104
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton Command87 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Torneo 2vs2"
         Height          =   375
         Left            =   240
         MaskColor       =   &H000000C0&
         TabIndex        =   103
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command57 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Hacerlo Neutro"
      Height          =   255
      Left            =   4560
      Style           =   1  'Graphical
      TabIndex        =   97
      Top             =   7320
      Width           =   2175
   End
   Begin VB.CommandButton Command85 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Activar 2 vs 2"
      Height          =   255
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   96
      Top             =   7560
      UseMaskColor    =   -1  'True
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFF80&
      Caption         =   "Controles de GM"
      ForeColor       =   &H00C00000&
      Height          =   3255
      Left            =   0
      TabIndex        =   73
      Top             =   0
      Width           =   4815
      Begin VB.Frame Frame8 
         BackColor       =   &H00FFFF80&
         Caption         =   "General"
         ForeColor       =   &H00C00000&
         Height          =   1095
         Left            =   120
         TabIndex        =   86
         Top             =   240
         Width           =   4455
         Begin VB.CommandButton Command35 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Invisible"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   93
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton Command36 
            BackColor       =   &H00C0FFFF&
            Caption         =   "GMs Online"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   92
            Top             =   480
            Width           =   1815
         End
         Begin VB.CommandButton Command37 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Usuarios Online"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   91
            Top             =   720
            Width           =   1815
         End
         Begin VB.TextBox Text15 
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   3480
            TabIndex        =   90
            Text            =   "Número"
            Top             =   480
            Width           =   735
         End
         Begin VB.CommandButton Command53 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Crear Item"
            Height          =   255
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   89
            Top             =   480
            Width           =   1095
         End
         Begin VB.CommandButton Command54 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Destruir Item"
            Height          =   255
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   720
            Width           =   1815
         End
         Begin VB.CommandButton Command51 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Destruir todos los Items"
            Height          =   255
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FFFF80&
         Caption         =   "Teleports"
         ForeColor       =   &H00C00000&
         Height          =   1695
         Left            =   120
         TabIndex        =   80
         Top             =   1440
         Width           =   2055
         Begin VB.TextBox Text4 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   120
            TabIndex        =   85
            Text            =   "Mapa"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox Text5 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   720
            TabIndex        =   84
            Text            =   "X"
            Top             =   360
            Width           =   495
         End
         Begin VB.TextBox Text6 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   1320
            TabIndex        =   83
            Text            =   "Y"
            Top             =   360
            Width           =   495
         End
         Begin VB.CommandButton Command33 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Crear Teleport"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   82
            Top             =   840
            Width           =   1695
         End
         Begin VB.CommandButton Command34 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Eliminar Teleport"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   81
            Top             =   1080
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FFFF80&
         Caption         =   "NPCs"
         ForeColor       =   &H00C00000&
         Height          =   1695
         Left            =   2280
         TabIndex        =   75
         Top             =   1440
         Width           =   2295
         Begin VB.CommandButton Command39 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Sumonear Con ReSpawm"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   79
            Top             =   600
            Width           =   2055
         End
         Begin VB.TextBox Text7 
            Alignment       =   2  'Center
            ForeColor       =   &H00800000&
            Height          =   285
            Left            =   120
            TabIndex        =   78
            Text            =   "Numero de NPC"
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton Command40 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Sumonear Sin ReSpawm"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   77
            Top             =   840
            Width           =   2055
         End
         Begin VB.CommandButton Command41 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Lista de Npcs"
            Height          =   255
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   1080
            Width           =   2055
         End
      End
      Begin VB.CommandButton Command48 
         Caption         =   "Borrar todos los NPCS"
         Height          =   255
         Left            =   2400
         TabIndex        =   74
         Top             =   2760
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFF80&
      Caption         =   "Mensajes"
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   2160
      TabIndex        =   68
      Top             =   3240
      Width           =   4815
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   240
         TabIndex        =   72
         Text            =   "Escribir el Mensaje"
         Top             =   240
         Width           =   4215
      End
      Begin VB.CommandButton Command24 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Enviar Mensaje Al Staff"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   600
         Width           =   4215
      End
      Begin VB.CommandButton Command23 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Enviar Mensaje en Cartel"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   840
         Width           =   4215
      End
      Begin VB.CommandButton Command22 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Enviar Mensaje en Consola"
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   1080
         Width           =   4215
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFF80&
      Caption         =   "Varios"
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   0
      TabIndex        =   63
      Top             =   3240
      Width           =   2175
      Begin VB.CommandButton Command28 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Restringir Servidor"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   1200
         Width           =   1935
      End
      Begin VB.CommandButton Command27 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Hora"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   960
         Width           =   1935
      End
      Begin VB.CommandButton Command25 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Limpiar Mundo"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   65
         Top             =   480
         Width           =   1935
      End
      Begin VB.CommandButton Command20 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Hacer WorldSave"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFF80&
      Caption         =   "Editar Usuarios"
      ForeColor       =   &H00800000&
      Height          =   3015
      Left            =   0
      TabIndex        =   38
      Top             =   4920
      Width           =   6975
      Begin VB.CommandButton Command59 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Hacerlo Azul"
         Height          =   255
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   99
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CommandButton Command58 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Hacerlo Rojo"
         Height          =   255
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   98
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CommandButton Command19 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Info"
         Height          =   255
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   2400
         Width           =   2295
      End
      Begin VB.CommandButton Command18 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Inventario"
         Height          =   255
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   2160
         Width           =   2295
      End
      Begin VB.CommandButton Command17 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ver skills de usuario"
         Height          =   255
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2295
      End
      Begin VB.CommandButton Command16 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Unban"
         Height          =   255
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton Command13 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Neutrales Matados"
         Height          =   255
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   58
         TabStop         =   0   'False
         Top             =   1200
         Width           =   2175
      End
      Begin VB.CommandButton Command12 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Matar"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   2400
         Width           =   2175
      End
      Begin VB.CommandButton Command11 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Revivir"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   2160
         Width           =   2175
      End
      Begin VB.CommandButton Command10 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Echar"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   1920
         Width           =   2175
      End
      Begin VB.CommandButton Command9 
         BackColor       =   &H00C0FFFF&
         Caption         =   "+"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   54
         TabStop         =   0   'False
         Top             =   1680
         Width           =   2175
      End
      Begin VB.CommandButton Command8 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ir al Usuario"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton Command7 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Aceptar en el Concilio de Arghal"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   1200
         Width           =   4455
      End
      Begin VB.CommandButton Command6 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Aceptar en el honorable Consejo de Banderbill"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   51
         TabStop         =   0   'False
         Top             =   960
         Width           =   4455
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Criminales Matados"
         Height          =   255
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   50
         TabStop         =   0   'False
         Top             =   960
         Width           =   2175
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ciudadanos Matados"
         Height          =   255
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   720
         Width           =   2175
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Oro"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   720
         Width           =   2175
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2400
         TabIndex        =   47
         Text            =   "Numero o Cantidad"
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   46
         Text            =   "Nick del PJ"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command44 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Banear IP"
         Height          =   255
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   1440
         Width           =   2175
      End
      Begin VB.CommandButton Command45 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Expulsar"
         Height          =   255
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   44
         Top             =   1440
         Width           =   2295
      End
      Begin VB.CommandButton Command49 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cabeza"
         Height          =   255
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text8 
         Alignment       =   2  'Center
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   4680
         TabIndex        =   42
         Text            =   "Escriba la clase"
         Top             =   360
         Width           =   2175
      End
      Begin VB.CommandButton Command63 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Morirse a sí mismo"
         Height          =   255
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   1920
         Width           =   2295
      End
      Begin VB.CommandButton Command64 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Activar Deathmatch"
         Height          =   255
         Left            =   4560
         Style           =   1  'Graphical
         TabIndex        =   40
         Top             =   2640
         Width           =   2175
      End
      Begin VB.CommandButton Command65 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Activar SOPORTE"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   39
         Top             =   2640
         Width           =   2295
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFF80&
      Caption         =   "Mapas"
      ForeColor       =   &H00C00000&
      Height          =   2415
      Left            =   4920
      TabIndex        =   15
      Top             =   0
      Width           =   3735
      Begin VB.CommandButton Command43 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ir"
         Height          =   255
         Index           =   1
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   100
         Top             =   1080
         Width           =   375
      End
      Begin VB.CommandButton Command69 
         BackColor       =   &H00C0FFFF&
         Caption         =   " Seguro"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   95
         Top             =   2040
         Width           =   1575
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ir"
         Height          =   255
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   960
         Width           =   375
      End
      Begin VB.CommandButton Command32 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ir"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   600
         Width           =   375
      End
      Begin VB.CommandButton Command31 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ir"
         Height          =   255
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   720
         Width           =   375
      End
      Begin VB.CommandButton Command30 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ir"
         Height          =   255
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   480
         Width           =   375
      End
      Begin VB.CommandButton Command29 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ir"
         Height          =   255
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   240
         Width           =   375
      End
      Begin VB.CommandButton Command14 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ir"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   840
         Width           =   375
      End
      Begin VB.CommandButton Command15 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ir"
         Height          =   255
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1200
         Width           =   375
      End
      Begin VB.CommandButton Command26 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ir"
         Height          =   255
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   1440
         Width           =   375
      End
      Begin VB.CommandButton Command42 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ir"
         Height          =   255
         Left            =   1320
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command43 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ir"
         Height          =   255
         Index           =   0
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   1680
         Width           =   375
      End
      Begin VB.CommandButton Command74 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ir"
         Height          =   255
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Prision"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   1800
         TabIndex        =   101
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sala 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   37
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Deathmatch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   1800
         TabIndex        =   36
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sala 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   35
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sala 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   34
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ullathorpe"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Batalla Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   1800
         TabIndex        =   32
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sala 4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   31
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Torneos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   30
         Top             =   1440
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Salas 2vs2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   29
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Sala Gms"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   1800
         TabIndex        =   28
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "/Quest"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   1800
         TabIndex        =   27
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command73 
      Caption         =   "Ir"
      Height          =   255
      Left            =   7920
      TabIndex        =   14
      Top             =   240
      Width           =   375
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00FFFF80&
      Caption         =   "Quests"
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   4920
      TabIndex        =   12
      Top             =   2400
      Width           =   2055
      Begin VB.CommandButton Command75 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Activar/Desactivar"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame10 
      BackColor       =   &H00FFFF80&
      Caption         =   "Torneos"
      ForeColor       =   &H00C00000&
      Height          =   2775
      Left            =   6960
      TabIndex        =   4
      Top             =   2400
      Width           =   1695
      Begin VB.CommandButton Command76 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Activar/Desactivar"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command77 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ver peticiones"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox Text9 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "Cuenta Regresiva"
         Top             =   940
         Width           =   1400
      End
      Begin VB.CommandButton Command78 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Cuenta Regresiva"
         Height          =   315
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox Text10 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   7
         Text            =   "Nick"
         Top             =   1680
         Width           =   1455
      End
      Begin VB.CommandButton Command79 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Sumonear"
         Height          =   255
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   2040
         Width           =   1455
      End
      Begin VB.CommandButton Command80 
         BackColor       =   &H00C0FFFF&
         Caption         =   "LLevar a Ulla"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   2280
         Width           =   1455
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00FFFF80&
      Caption         =   "Más cosas"
      ForeColor       =   &H00C00000&
      Height          =   1455
      Left            =   6960
      TabIndex        =   0
      Top             =   5160
      Width           =   1695
      Begin VB.CommandButton Command84 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Ignorar NPCs"
         Height          =   255
         Left            =   120
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   94
         Top             =   840
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.TextBox Text11 
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Text            =   "Nick"
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Command81 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Proteger User"
         Height          =   255
         Left            =   120
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   600
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
      Begin VB.CommandButton Command83 
         BackColor       =   &H00C0FFFF&
         Caption         =   "Vigilar"
         Height          =   255
         Left            =   120
         MaskColor       =   &H000000C0&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1080
         UseMaskColor    =   -1  'True
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmComandosGM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
    Call SendData("/telep yo 5 50 50")
End Sub


Private Sub Command10_Click()
    Call SendData("/ECHAR" & " " & Text1.Text)
End Sub

Private Sub Command11_Click()
    Call SendData("/REVIVIR" & " " & Text1.Text)
End Sub

Private Sub Command12_Click()
    Call SendData("/KILL" & " " & Text1.Text)
End Sub

Private Sub Command13_Click()
    Call SendData("/MOD" & " " & Text1.Text & " " & "Neu" & " " & Text2.Text)
End Sub

Private Sub Command14_Click()
    Call SendData("/telep yo 11 50 50")
End Sub

Private Sub Command15_Click()
    Call SendData("/telep yo 6 50 50")
End Sub

Private Sub Command16_Click()
    If MsgBox("Esta seguro que desea removerle el Ban a dicho pj?", vbYesNo) = vbYes Then
        Call SendData("/UNBAN " & Text1.Text)
    End If
End Sub

Private Sub Command17_Click()
    Call SendData("/SKILLS")
End Sub

Private Sub Command18_Click()
    Call SendData("/INV")
End Sub

Private Sub Command2_Click()
    Call SendData("/MOD" & " " & Text1.Text & " " & "oro" & " " & Text2.Text)
End Sub

Private Sub Command20_Click()
    Call SendData("/DOBACKUP")
End Sub


Private Sub Command22_Click()
    Call SendData("/RMSG" & " " & Text3.Text)
End Sub

Private Sub Command23_Click()
    Call SendData("/SMSG" & " " & Text3.Text)
End Sub

Private Sub Command24_Click()
    Call SendData("/STAFF" & " " & Text3.Text)
End Sub

Private Sub Command25_Click()
    Call SendData("/LIMPIARMUNDO")
End Sub

Private Sub Command26_Click()
    Call SendData("/telep yo 14 50 50")
End Sub

Private Sub Command27_Click()
    Call SendData("/HORA")
End Sub

Private Sub Command28_Click()
    Call SendData("/RESTRINGIR")
End Sub

Private Sub Command29_Click()
    Call SendData("/telep yo 1 50 50")
End Sub
Private Sub Command30_Click()
    Call SendData("/telep yo 3 50 50")
End Sub

Private Sub Command31_Click()
    Call SendData("/telep yo 4 50 50")
End Sub

Private Sub Command32_Click()
    Call SendData("/telep yo 12 50 50")
End Sub

Private Sub Command33_Click()
    Call SendData("/CT" & " " & Text4.Text & " " & Text5.Text & " " & Text6.Text)
End Sub

Private Sub Command34_Click()
    Call SendData("/DT")
End Sub

Private Sub Command35_Click()
    Call SendData("/INVISIBLE")
End Sub

Private Sub Command36_Click()
    Call SendData("/ONLINEGM")
End Sub
Private Sub Command37_Click()
    Call SendData("/ONLINE")
End Sub


Private Sub Command39_Click()
    Call SendData("/RACC" & " " & Text7.Text)
End Sub

Private Sub Command4_Click()
    Call SendData("/MOD" & " " & Text1.Text & " " & "Cri" & " " & Text2.Text)
End Sub

Private Sub Command40_Click()
    Call SendData("/ACC" & " " & Text7.Text)
End Sub

Private Sub Command41_Click()
    Call SendData("/CC")
End Sub

Private Sub Command42_Click()
    Call SendData("/telep yo 7 50 50")
End Sub


Private Sub Command43_Click(Index As Integer)
    Call SendData("/telep yo 17 50 50")
End Sub

Private Sub Command44_Click()
    Call SendData("/BANIP" & " " & Text1.Text)
End Sub

Private Sub Command45_Click()    'Este si lo TIENEN
    Call SendData("/ECHARCONSE" & " " & Text1.Text)    'Este si lo TIENEN
End Sub    'Este si lo TIENEN

Private Sub Command49_Click()
    Call SendData("/MOD" & " " & Text1.Text & " " & "head" & " " & Text2.Text)
End Sub

Private Sub Command5_Click()
    Call SendData("/MOD" & " " & Text1.Text & " " & "Ciu" & " " & Text2.Text)
End Sub

Private Sub Command51_Click()
    Call SendData("/MASSDEST")
End Sub


Private Sub Command53_Click()
    Call SendData("/ITEM" & " " & Text15.Text)
End Sub

Private Sub Command54_Click()
    Call SendData("/DEST")
End Sub

Private Sub Command57_Click()
    Call SendData("/MOD" & " " & Text1.Text & " " & "bando" & " " & "0")
End Sub

Private Sub Command58_Click()
    Call SendData("/MOD" & " " & Text1.Text & " " & "bando" & " " & "2")
End Sub

Private Sub Command59_Click()
    Call SendData("/MOD" & " " & Text1.Text & " " & "bando" & " " & "1")
End Sub

Private Sub Command6_Click()    'Este si lo TIENEN
    Call SendData("/ACEPTCONSE" & " " & Text1.Text)    'Este si lo TIENEN
End Sub    'Este si lo TIENEN


Private Sub Command63_Click()
    Call SendData("/DIE")
End Sub


Private Sub Command64_Click()
    Call SendData("/SISTEMADEATHMATCH")
End Sub

Private Sub Command65_Click()
    Call SendData("/SISTEMASOPORTE")
End Sub

Private Sub Command69_Click()
    Call SendData("/SEGURO")
End Sub

Private Sub Command7_Click()    'Este si lo TIENEN
    Call SendData("/ACEPTCONSECAOS" & " " & Text1.Text)    'Este si lo TIENEN
End Sub    'Este si lo TIENEN

Private Sub Command73_Click()
    Call SendData("/GO 22")
End Sub

Private Sub Command74_Click()
    Call SendData("/telep yo 10 50 50")
End Sub

Private Sub Command75_Click()
    Call SendData("/MODOQUEST")
End Sub

Private Sub Command76_Click()
    Call SendData("/TORNEO")
End Sub

Private Sub Command77_Click()
    Call SendData("/SHOW TORNEO")
End Sub

Private Sub Command78_Click()
    Call SendData("/CUENTA" & " " & Text9.Text)
End Sub

Private Sub Command79_Click()
    Call SendData("/SUM" & " " & Text10.Text)
End Sub

Private Sub Command8_Click()
    Call SendData("/IRA" & " " & Text1.Text)
End Sub

Private Sub Command80_Click()
    Call SendData("/TELEP" & " " & Text10.Text & "1 50 50")
End Sub

Private Sub Command81_Click()
    Call SendData("/PRO" & " " & Text11.Text)
End Sub

Private Sub Command83_Click()
    Call SendData("/VIGILAR" & " " & Text11.Text)
End Sub

Private Sub Command85_Click()
    Call SendData("/SISTEMARETOS")
End Sub

Private Sub Command87_Click()
    Torneo2.Show
End Sub

Private Sub Command88_Click()
    Torneo1.Show
End Sub

Private Sub Command9_Click()
    Call SendData("/SUM" & " " & Text1.Text)
End Sub

