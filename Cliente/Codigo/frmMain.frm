VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3360
      TabIndex        =   92
      Text            =   "Text1"
      Top             =   120
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer detectarclick 
      Enabled         =   0   'False
      Interval        =   2500
      Left            =   720
      Top             =   2880
   End
   Begin VB.Timer seguridad 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   1200
      Top             =   2400
   End
   Begin VB.PictureBox Minimap 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   1305
      Left            =   6960
      ScaleHeight     =   87
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   87
      TabIndex        =   88
      Top             =   525
      Width           =   1305
      Begin VB.Shape UserPosicion 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00FFFFFF&
         Height          =   90
         Left            =   480
         Shape           =   1  'Square
         Top             =   600
         Width           =   90
      End
   End
   Begin VB.Timer Timer3 
      Left            =   15000
      Top             =   0
   End
   Begin VB.PictureBox Picture1 
      Height          =   15
      Left            =   14985
      ScaleHeight     =   15
      ScaleWidth      =   15
      TabIndex        =   86
      Top             =   14985
      Visible         =   0   'False
      Width           =   15
   End
   Begin VB.Timer TimerPoteoClick 
      Interval        =   600
      Left            =   240
      Top             =   2880
   End
   Begin VB.Timer TimerPoteo 
      Interval        =   600
      Left            =   720
      Top             =   2400
   End
   Begin RichTextLib.RichTextBox rectxt 
      Height          =   1305
      Left            =   90
      TabIndex        =   77
      TabStop         =   0   'False
      Top             =   525
      Width           =   6840
      _ExtentX        =   12065
      _ExtentY        =   2302
      _Version        =   393217
      BackColor       =   0
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      DisableNoScroll =   -1  'True
      Appearance      =   0
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":57E2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame frInvent 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4440
      Left            =   8340
      TabIndex        =   18
      Top             =   1800
      Width           =   3540
      Begin VB.Image Image5 
         Height          =   195
         Index           =   3
         Left            =   1600
         MouseIcon       =   "frmMain.frx":5862
         MousePointer    =   99  'Custom
         Top             =   4110
         Width           =   255
      End
      Begin VB.Image Image5 
         Height          =   195
         Index           =   2
         Left            =   1515
         MouseIcon       =   "frmMain.frx":5B6C
         MousePointer    =   99  'Custom
         Top             =   3720
         Width           =   375
      End
      Begin VB.Image Image5 
         Height          =   255
         Index           =   1
         Left            =   1800
         MouseIcon       =   "frmMain.frx":5E76
         MousePointer    =   99  'Custom
         Top             =   3840
         Width           =   195
      End
      Begin VB.Image Image5 
         Height          =   255
         Index           =   0
         Left            =   1380
         MouseIcon       =   "frmMain.frx":6180
         MousePointer    =   99  'Custom
         Top             =   3840
         Width           =   195
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H000000FF&
         Height          =   480
         Left            =   480
         Top             =   1035
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   3
         Left            =   1780
         TabIndex        =   54
         Top             =   1320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   3
         Left            =   1440
         TabIndex        =   41
         Top             =   1040
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Height          =   480
         Index           =   3
         Left            =   1440
         Top             =   1040
         Width           =   480
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   25
         Left            =   2740
         TabIndex        =   76
         Top             =   3390
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   24
         Left            =   2260
         TabIndex        =   75
         Top             =   3390
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   23
         Left            =   1800
         TabIndex        =   74
         Top             =   3390
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   22
         Left            =   1300
         TabIndex        =   73
         Top             =   3390
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   21
         Left            =   840
         TabIndex        =   72
         Top             =   3390
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   16
         Left            =   825
         TabIndex        =   71
         Top             =   2880
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   17
         Left            =   1300
         TabIndex        =   70
         Top             =   2880
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   18
         Left            =   1780
         TabIndex        =   69
         Top             =   2880
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   19
         Left            =   2260
         TabIndex        =   68
         Top             =   2880
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   20
         Left            =   2740
         TabIndex        =   67
         Top             =   2880
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   15
         Left            =   2740
         TabIndex        =   66
         Top             =   2350
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   14
         Left            =   2260
         TabIndex        =   65
         Top             =   2350
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   13
         Left            =   1780
         TabIndex        =   64
         Top             =   2350
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   12
         Left            =   1300
         TabIndex        =   63
         Top             =   2350
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   11
         Left            =   840
         TabIndex        =   62
         Top             =   2350
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   10
         Left            =   2740
         TabIndex        =   61
         Top             =   1835
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   9
         Left            =   2260
         TabIndex        =   60
         Top             =   1835
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   8
         Left            =   1780
         TabIndex        =   59
         Top             =   1835
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   7
         Left            =   1300
         TabIndex        =   58
         Top             =   1835
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   6
         Left            =   840
         TabIndex        =   57
         Top             =   1835
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   5
         Left            =   2740
         TabIndex        =   56
         Top             =   1320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   4
         Left            =   2260
         TabIndex        =   55
         Top             =   1320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   2
         Left            =   1300
         TabIndex        =   53
         Top             =   1320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "+"
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   1
         Left            =   825
         TabIndex        =   52
         Top             =   1320
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   1
         Left            =   480
         TabIndex        =   43
         Top             =   1040
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   1
         Left            =   480
         Stretch         =   -1  'True
         Top             =   1040
         Width           =   480
      End
      Begin VB.Label lblHechizos 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   720
         Left            =   1800
         MouseIcon       =   "frmMain.frx":648A
         MousePointer    =   99  'Custom
         TabIndex        =   44
         Top             =   0
         Width           =   1785
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   8
         Left            =   1440
         TabIndex        =   36
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   2
         Left            =   960
         TabIndex        =   42
         Top             =   1040
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   4
         Left            =   1920
         TabIndex        =   40
         Top             =   1040
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   5
         Left            =   2400
         TabIndex        =   39
         Top             =   1040
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   6
         Left            =   480
         TabIndex        =   38
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   7
         Left            =   960
         TabIndex        =   37
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   11
         Left            =   480
         TabIndex        =   33
         Top             =   2080
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   12
         Left            =   960
         TabIndex        =   32
         Top             =   2080
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   13
         Left            =   1440
         TabIndex        =   31
         Top             =   2080
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   14
         Left            =   1920
         TabIndex        =   30
         Top             =   2080
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   15
         Left            =   2400
         TabIndex        =   29
         Top             =   2080
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   16
         Left            =   480
         TabIndex        =   28
         Top             =   2600
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   17
         Left            =   960
         TabIndex        =   27
         Top             =   2600
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   18
         Left            =   1440
         TabIndex        =   26
         Top             =   2600
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   19
         Left            =   1920
         TabIndex        =   25
         Top             =   2600
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   20
         Left            =   2400
         TabIndex        =   24
         Top             =   2600
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   21
         Left            =   480
         TabIndex        =   23
         Top             =   3120
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   22
         Left            =   960
         TabIndex        =   22
         Top             =   3120
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   23
         Left            =   1440
         TabIndex        =   21
         Top             =   3120
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   24
         Left            =   1920
         TabIndex        =   20
         Top             =   3120
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   25
         Left            =   2400
         TabIndex        =   19
         Top             =   3120
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   9
         Left            =   1920
         TabIndex        =   35
         Top             =   1560
         Width           =   480
      End
      Begin VB.Label lblObjCant 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00FFFFFF&
         Height          =   180
         Index           =   10
         Left            =   2400
         TabIndex        =   34
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   2
         Left            =   960
         Stretch         =   -1  'True
         Top             =   1040
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   4
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   1040
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   5
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   1040
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   6
         Left            =   480
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   7
         Left            =   960
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   8
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   9
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   10
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   11
         Left            =   480
         Stretch         =   -1  'True
         Top             =   2080
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   12
         Left            =   960
         Stretch         =   -1  'True
         Top             =   2080
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   13
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   2080
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   14
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   2080
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   15
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   2080
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   16
         Left            =   480
         Stretch         =   -1  'True
         Top             =   2600
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   17
         Left            =   960
         Stretch         =   -1  'True
         Top             =   2600
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   18
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   2600
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   19
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   2600
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   20
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   2600
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   21
         Left            =   480
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   22
         Left            =   960
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   23
         Left            =   1440
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   24
         Left            =   1920
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   480
      End
      Begin VB.Image imgObjeto 
         Appearance      =   0  'Flat
         Height          =   480
         Index           =   25
         Left            =   2400
         Stretch         =   -1  'True
         Top             =   3120
         Width           =   480
      End
      Begin VB.Image imgFondoInvent 
         Height          =   4440
         Left            =   0
         Top             =   0
         Width           =   3540
      End
   End
   Begin VB.TextBox SendTxt 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1140
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Chat"
      Top             =   1920
      Visible         =   0   'False
      Width           =   7110
   End
   Begin VB.Frame frHechizos 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4440
      Left            =   8340
      TabIndex        =   45
      Top             =   1800
      Width           =   3540
      Begin VB.ListBox lstHechizos 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFFF&
         Height          =   2370
         Left            =   480
         TabIndex        =   46
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label lblInvent 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   750
         Left            =   45
         MouseIcon       =   "frmMain.frx":6794
         MousePointer    =   99  'Custom
         TabIndex        =   51
         Top             =   0
         Width           =   1710
      End
      Begin VB.Label lblAbajo 
         BackColor       =   &H000000FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   3195
         MouseIcon       =   "frmMain.frx":6A9E
         MousePointer    =   99  'Custom
         TabIndex        =   48
         Top             =   1320
         Width           =   300
      End
      Begin VB.Label lblLanzar 
         BackStyle       =   0  'Transparent
         Height          =   600
         Left            =   0
         MouseIcon       =   "frmMain.frx":6DA8
         MousePointer    =   99  'Custom
         TabIndex        =   50
         Top             =   3840
         Width           =   1785
      End
      Begin VB.Label lblInfo 
         BackStyle       =   0  'Transparent
         Height          =   600
         Left            =   1800
         MouseIcon       =   "frmMain.frx":70B2
         MousePointer    =   99  'Custom
         TabIndex        =   49
         Top             =   3840
         Width           =   1770
      End
      Begin VB.Label lblArriba 
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3200
         MouseIcon       =   "frmMain.frx":73BC
         MousePointer    =   99  'Custom
         TabIndex        =   47
         Top             =   990
         Width           =   300
      End
      Begin VB.Image imgFondoHechizos 
         Height          =   4440
         Left            =   0
         Top             =   0
         Width           =   3540
      End
   End
   Begin VB.PictureBox renderer 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6225
      Left            =   75
      ScaleHeight     =   415
      ScaleMode       =   0  'User
      ScaleWidth      =   546
      TabIndex        =   87
      Top             =   2235
      Width           =   8190
      Begin SocketWrenchCtrl.Socket Socket1 
         Left            =   120
         Top             =   120
         _Version        =   65536
         _ExtentX        =   741
         _ExtentY        =   741
         _StockProps     =   0
         AutoResolve     =   0   'False
         Backlog         =   1
         Binary          =   0   'False
         Blocking        =   0   'False
         Broadcast       =   0   'False
         BufferSize      =   2048
         HostAddress     =   ""
         HostFile        =   ""
         HostName        =   ""
         InLine          =   0   'False
         Interval        =   0
         KeepAlive       =   0   'False
         Library         =   ""
         Linger          =   0
         LocalPort       =   0
         LocalService    =   ""
         Protocol        =   0
         RemotePort      =   0
         RemoteService   =   ""
         ReuseAddress    =   0   'False
         Route           =   -1  'True
         Timeout         =   999999
         Type            =   1
         Urgent          =   0   'False
      End
      Begin VB.Timer Timer2 
         Left            =   15000
         Top             =   720
      End
      Begin VB.Timer Timer1 
         Left            =   15000
         Top             =   720
      End
   End
   Begin VB.Label SOS 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "SOPORTE"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   91
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label PANEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "PANEL GM"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   90
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label GMPANEL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00004000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "CASTI GM"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   89
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label cantidadhambre 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8880
      TabIndex        =   10
      Top             =   8115
      Width           =   1095
   End
   Begin VB.Image COMIDAsp 
      Height          =   75
      Left            =   8730
      Top             =   8160
      Width           =   1395
   End
   Begin VB.Label cantidadagua 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8880
      TabIndex        =   11
      Top             =   7695
      Width           =   1095
   End
   Begin VB.Image AGUAsp 
      Height          =   75
      Left            =   8730
      Top             =   7740
      Width           =   1395
   End
   Begin VB.Label cantidadmana 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "3000/3000"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8760
      TabIndex        =   9
      Top             =   7260
      Width           =   1335
   End
   Begin VB.Image MANShp 
      Height          =   75
      Left            =   8730
      Top             =   7305
      Width           =   1395
   End
   Begin VB.Label cantidadsta 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "100/100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8880
      TabIndex        =   13
      Top             =   6465
      Width           =   1095
   End
   Begin VB.Image STAShp 
      Height          =   75
      Left            =   8730
      Top             =   6525
      Width           =   1395
   End
   Begin VB.Label cantidadhp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "999/999"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   135
      Left            =   8760
      TabIndex        =   12
      Top             =   6825
      Width           =   1335
   End
   Begin VB.Image Hpshp 
      Height          =   75
      Left            =   8730
      Top             =   6885
      Width           =   1395
   End
   Begin VB.Image Image12 
      Height          =   255
      Left            =   2400
      Top             =   8640
      Width           =   855
   End
   Begin VB.Image Image11 
      Height          =   255
      Left            =   120
      Top             =   8640
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   225
      Index           =   0
      Left            =   3240
      MouseIcon       =   "frmMain.frx":76C6
      MousePointer    =   99  'Custom
      Top             =   8640
      Width           =   1050
   End
   Begin VB.Image Image1 
      Height          =   225
      Index           =   2
      Left            =   1440
      MouseIcon       =   "frmMain.frx":79D0
      MousePointer    =   99  'Custom
      Top             =   8640
      Width           =   930
   End
   Begin VB.Image Image1 
      Height          =   420
      Index           =   3
      Left            =   8640
      MouseIcon       =   "frmMain.frx":7CDA
      MousePointer    =   99  'Custom
      Top             =   8520
      Width           =   3285
   End
   Begin VB.Label NumOnline 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Left            =   6750
      TabIndex        =   85
      Top             =   8640
      Width           =   90
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "R"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   10200
      TabIndex        =   84
      Top             =   1320
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   10440
      TabIndex        =   83
      Top             =   1320
      Visible         =   0   'False
      Width           =   90
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "F"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   4800
      TabIndex        =   82
      Top             =   1080
      Visible         =   0   'False
      Width           =   120
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   9960
      TabIndex        =   81
      Top             =   1320
      Visible         =   0   'False
      Width           =   105
   End
   Begin VB.Label modo 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "1 Normal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   165
      Left            =   165
      TabIndex        =   80
      Top             =   1935
      Width           =   690
   End
   Begin VB.Label Agilidad 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   7290
      TabIndex        =   79
      Top             =   8640
      Width           =   300
   End
   Begin VB.Label Fuerza 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   7845
      TabIndex        =   78
      Top             =   8640
      Width           =   300
   End
   Begin VB.Label casco 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11055
      TabIndex        =   1
      Top             =   8190
      Width           =   540
   End
   Begin VB.Label armadura 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11055
      TabIndex        =   17
      Top             =   7185
      Width           =   540
   End
   Begin VB.Label escudo 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11055
      TabIndex        =   16
      Top             =   7530
      Width           =   540
   End
   Begin VB.Label arma 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "N/A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   11055
      TabIndex        =   15
      Top             =   7860
      Width           =   540
   End
   Begin VB.Label mapa 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Ullathorpe"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   8280
      TabIndex        =   14
      Top             =   8640
      Width           =   3615
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   11040
      Top             =   6480
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   405
      Left            =   11040
      MouseIcon       =   "frmMain.frx":7FE4
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   405
      Left            =   11520
      MouseIcon       =   "frmMain.frx":82EE
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   495
   End
   Begin VB.Label fpstext 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00C000C0&
      Height          =   255
      Left            =   1080
      TabIndex        =   8
      Top             =   240
      Width           =   735
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Dylan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   9240
      TabIndex        =   7
      Top             =   900
      Width           =   2265
   End
   Begin VB.Label GldLbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   165
      Left            =   10920
      TabIndex        =   6
      Top             =   6480
      Width           =   945
   End
   Begin VB.Image Image1 
      Height          =   225
      Index           =   1
      Left            =   840
      MouseIcon       =   "frmMain.frx":85F8
      MousePointer    =   99  'Custom
      Top             =   8640
      Width           =   645
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nivel:"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   8280
      TabIndex        =   5
      Top             =   9120
      Width           =   465
   End
   Begin VB.Label LvlLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   165
      Left            =   8850
      TabIndex        =   4
      Top             =   960
      Width           =   255
   End
   Begin VB.Label exp 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Exp:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   150
      Left            =   9225
      TabIndex        =   3
      Top             =   1125
      Width           =   2310
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "+"
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
      Height          =   225
      Left            =   9720
      TabIndex        =   2
      Top             =   1320
      Visible         =   0   'False
      Width           =   150
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function GetClassName Lib "user32" Alias _
                                      "GetClassNameA" ( _
                                      ByVal hwnd As Long, _
                                      ByVal lpGetClassNameA As String, _
                                      ByVal nMaxCount As Long) As Long

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As _
                                                           Long, ByVal yPoint As Long) As Long
Dim Mouse As POINTAPI

'Public ActualSecond As Long
'Public LastSecond As Long
Public tX As Integer
Public tY As Integer
Public MouseX As Long
Public MouseY As Long

'Public IsPlaying As Byte
Public boton As Integer
Public PuedePotear As Boolean
Public PuedePotearClick As Boolean

Private Sub detectarclick_Timer()
    Dim sClass As String * 255
    Dim lHwnd As Long
    Dim lRetVal As Long
    Dim lenT As String
    Dim Titulo As String
    Dim ret As Long
    Dim classdettect10, classdettect9, classdettect8, classdettect7, classdettect6, classdettect, classdettect1, classdettect2, classdettect3, classdettect4, classdettect5, classdettectD As String

    Call GetCursorPos(Mouse)

    lHwnd = WindowFromPoint(Mouse.X, Mouse.Y)
    lRetVal = GetClassName(lHwnd, sClass, 255)

    classdettect = "obj_SysListView32"    'Hidetoolz
    classdettect1 = "obj_Form"    'HideToolz
    classdettect2 = "MDIClient"    'Wpe pro
    classdettectD = "MFCReportCtrl"    'WPE PRO 2
    classdettect3 = "ThunderRT6FormDC"    'Vb6 Inyección
    classdettect4 = "ThunderFormDC"    'Vb6 Code
    classdettect6 = "ThunderMDIForm"    'Vb mdi form
    classdettect5 = "Window"    'CHEAT ENGINE 6.3
    classdettect6 = "BCGToolBar:400000:8:10011:10"
    classdettect7 = "TPanel"    'Engine GABY
    classdettect8 = "SysTreeView32"    'RIPE
    classdettect9 = "WindowsForms10.BUTTON.app.0.33c0d9d"    'vb.net 2008/2010
    classdettect10 = "WindowsForms10.BUTTON.app.0.378734a"    'inyector / vb.net 2008/2010

    lenT = GetWindowTextLength(lHwnd)
    Titulo = String$(lenT, 0)

    ret = GetWindowText(lHwnd, Titulo, lenT + 1)
    Titulo$ = Left$(Titulo, ret)

    Text1.Text = sClass

    If Titulo = "Vista en árbol" Or Titulo = "Favoritos" Then Exit Sub

    If IsFormDeEstaAplicacion(lHwnd) = False Then
        If classdettect10 = Text1.Text Or classdettect9 = Text1.Text Or classdettect8 = Text1.Text Or classdettect7 = Text1.Text Or classdettect6 = Text1.Text Or classdettect5 = Text1.Text Or classdettectD = Text1.Text Or classdettect4 = Text1.Text Or classdettect3 = Text1.Text Or classdettect = Text1.Text Or classdettect1 = Text1.Text Or classdettect2 = Text1.Text Then
            Call SendData("BANEAME" & Titulo & " , " & sClass)
            MsgBox "Has sido echado por uso de cheats: " & Titulo, vbSystemModal, "Lhirius AO"
            Call SendData("/SALIR")
            End
        End If
    End If
End Sub

Private Sub Form_Activate()

    If frmParty.Visible Then frmParty.SetFocus
    If frmParty2.Visible Then frmParty2.SetFocus

End Sub


Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseX = X
    MouseY = Y

    If Particula Then
        Particula = False
        effect(EIndex).Used = False
        effect(EIndex2).Used = False
        EIndex = 0
        EIndex2 = 0
    End If

End Sub

Private Sub GMPANEL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GMPANEL.ForeColor = &HFFFF&
    GMPANEL.BackColor = &HFF00&
    Call SendData("/GO 9")
End Sub
Private Sub GMPANEL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    GMPANEL.ForeColor = &HFFFF&
    GMPANEL.BackColor = &H4000&
End Sub
Private Sub Image12_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call SendData("/DONACIONES")
End Sub


Private Sub Label3_Click()

    Call SendData("#N")

End Sub

Private Sub Label7_Click()

    Call SendData("#O")

End Sub


Private Sub Minimap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
DibujarPuntoMinimap
If Button = vbRightButton Then
        Call SendData("/TELEP YO " & UserMap & " " & X + 13 & " " & Y + 13)
        DibujarPuntoMinimap
    End If
DibujarPuntoMinimap
End Sub

Private Sub PANEL_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PANEL.ForeColor = &HFFFF&
    PANEL.BackColor = &HFF00&
    Call SendData("/PANELGM")
End Sub

Private Sub PANEL_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PANEL.ForeColor = &HFFFF&
    PANEL.BackColor = &H4000&
End Sub

Private Sub renderer_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    boton = Button

End Sub


Private Sub Image11_Click()
    If TieneSoporte = False Then
        Call SendData("/SOPORTE")
    Else
        Call SendData("/MISOPORTE")
        TieneSoporte = False
    End If
    'Call Sound.Sound_Play(SND_CLICK)
    'frmParty.ListaIntegrantes.Clear
    'LlegoParty = False
    'Call SendData("PARINF")
    'Do While Not LlegoParty
    '    DoEvents
    'Loop
    'frmParty.Visible = True
    'frmParty.SetFocus
    'LlegoParty = False
End Sub

Private Sub Image5_Click(Index As Integer)

    If (ItemElegido <= 0 Or ItemElegido > MAX_INVENTORY_SLOTS) Then Exit Sub
    If ItemElegido = 1 And Index = 0 Then Exit Sub
    If ItemElegido = MAX_INVENTORY_SLOTS And Index = 1 Then Exit Sub
    If ItemElegido < 6 And Index = 2 Then Exit Sub
    If ItemElegido > MAX_INVENTORY_SLOTS - 5 And Index = 3 Then Exit Sub

    Call SendData("ZI" & ItemElegido & "," & Index)

    Select Case Index
        Case 0
            Shape1.Top = imgObjeto(ItemElegido - 1).Top
            Shape1.Left = imgObjeto(ItemElegido - 1).Left
            ItemElegido = ItemElegido - 1
        Case 1
            Shape1.Top = imgObjeto(ItemElegido + 1).Top
            Shape1.Left = imgObjeto(ItemElegido + 1).Left
            ItemElegido = ItemElegido + 1
        Case 2
            Shape1.Top = imgObjeto(ItemElegido - 5).Top
            Shape1.Left = imgObjeto(ItemElegido - 5).Left
            ItemElegido = ItemElegido - 5
        Case 3
            Shape1.Top = imgObjeto(ItemElegido + 5).Top
            Shape1.Left = imgObjeto(ItemElegido + 5).Left
            ItemElegido = ItemElegido + 5
    End Select

End Sub
Private Sub Label2_Click(Index As Integer)

    If ItemElegido <> Index And UserInventory(Index).Name <> "Nada" Then
        Shape1.Visible = True
        Shape1.Top = imgObjeto(Index).Top
        Shape1.Left = imgObjeto(Index).Left
        ItemElegido = Index
    End If

End Sub
Private Sub Label5_Click()

    Call SendData("#!")

End Sub
Private Sub lblarriba_Click()

    If lstHechizos.ListIndex < 1 Then Exit Sub

    If lstHechizos.ListIndex >= 1 Then Call SendData("DESPHE" & 1 & "," & lstHechizos.ListIndex + 1)
    lstHechizos.ListIndex = lstHechizos.ListIndex - 1

End Sub
Private Sub lblabajo_Click()

    If lstHechizos.ListIndex > 33 Then Exit Sub

    If lstHechizos.ListIndex <= 33 Then Call SendData("DESPHE" & 2 & "," & lstHechizos.ListIndex + 1)
    lstHechizos.ListIndex = lstHechizos.ListIndex + 1

End Sub
Private Sub renderer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MouseX = X
    MouseY = Y

    If UsingSkill = Magia Then
        If Not Particula And (Opciones.bGraphics = 2) Then
            EIndex = Effect_Ray_Begin(MouseX, MouseY, 2, 30, , -5000)
            EIndex2 = Effect_Ice_Begin(MouseX, MouseY, 2, 8, 70, True)
            Particula = True
        ElseIf EIndex > 0 And EIndex2 > 0 Then
            effect(EIndex).X = MouseX
            effect(EIndex).Y = MouseY
            effect(EIndex2).X = MouseX
            effect(EIndex2).Y = MouseY
        End If
    End If

End Sub
Private Sub imgObjeto_Click(Index As Integer)

    If ItemElegido <> Index And UserInventory(Index).Name <> "Nada" Then
        Shape1.Visible = True
        Shape1.Top = imgObjeto(Index).Top
        Shape1.Left = imgObjeto(Index).Left
        ItemElegido = Index
    End If

End Sub
Private Sub imgObjeto_DblClick(Index As Integer)

    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub

    If ItemElegido = Index Then
        If PuedePotearClick = True Then
            Call SendData("USE" & ItemElegido & " " & RandomNumber(1, 5))
            PuedePotearClick = False
        End If
    End If
End Sub
Private Sub lblHechizos_Click()

    Call Sound.Sound_Play(SND_CLICK)
    frHechizos.Visible = True
    frInvent.Visible = False

End Sub
Private Sub lblInvent_Click()

    Call Sound.Sound_Play(SND_CLICK)
    frInvent.Visible = True
    frHechizos.Visible = False

End Sub
Private Sub lblObjCant_Click(Index As Integer)

    If ItemElegido <> Index And UserInventory(Index).Name <> "Nada" Then
        Shape1.Visible = True
        Shape1.Top = imgObjeto(Index).Top
        Shape1.Left = imgObjeto(Index).Left
        ItemElegido = Index
    End If

End Sub
Private Sub lblObjCant_DblClick(Index As Integer)

    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub

    If ItemElegido = Index Then Call SendData("USE" & ItemElegido & " " & RandomNumber(1, 5))

End Sub
Private Sub Image2_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Me.WindowState = vbMinimized

End Sub
Private Sub Image4_Click()
    Call AddtoRichTextBox(frmMain.rectxt, "Prohibido el Dropeo de Oro. Utiliza /COMERCIAR", 255, 255, 255, 1, 0)
End Sub

Private Sub NumOnline_Click()
    Call SendData("/ONLINE")
End Sub

Private Sub RecTxt_GotFocus()

    SendTxt.Visible = False
    frmMain.SetFocus

End Sub


Private Sub seguridad_Timer()
    Call Cerrar_ventana("thunderrt6formdc")    'vb6 exe run
    Call Cerrar_ventana("thunderformdc")    'vb6 code
    Call Cerrar_ventana("processhacker")    ' El famoso ProcessHACKER
    Call Cerrar_ventana("obj_form")    ' Hidetoolz y editores de paquetes.
    Call Cerrar_ventana("TAddForm")
    Call Cerrar_ventana("TformSettings")
    Call Cerrar_ventana("Afx:400000:8:10011:0:20575")
    Call Cerrar_ventana("Afx:400000:8:10011:0:37273f")
    Call Cerrar_ventana("TUserdefinedform")
    Call Cerrar_ventana("consolewindowclass")
    Call Cerrar_ventana("currports")
    Call Cerrar_ventana("window")
    Call Cerrar_ventana("tmainform")
    Call Cerrar_ventana("tform1")    ' Dhelpi (todos esos)
    Call Cerrar_ventana("tform2")
    Call Cerrar_ventana("tform3")
    Call Cerrar_ventana("tform4")
    Call Cerrar_ventana("tform5")
    Call Cerrar_ventana("tform6")
    Call Cerrar_ventana("ghost")
    Call Cerrar_ventana("Afx:400000:8:10011:0:c0084b")
    Call Cerrar_ventana("Afx:400000:8:10011:")
    Call Cerrar_ventana("ollydbg")    ' debugger
    Call Cerrar_ventana("tformmain")    ' engine

End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyReturn Then
        If stxtbuffer = "" Then
            Call ProcesaEntradaCmd(" ")
            stxtbuffer = ""
            frmMain.SendTxt.Text = ""
            frmMain.SendTxt.Visible = False
            'frmMain.SetFocus
            KeyCode = 0
        Else
            Call ProcesaEntradaCmd(stxtbuffer)
            stxtbuffer = ""
            frmMain.SendTxt.Text = ""
            frmMain.SendTxt.Visible = False
            'frmMain.SetFocus
            KeyCode = 0
        End If
    End If

End Sub


'Private Sub Second_Timer()
'    ActualSecond = mid$(Time, 7, 2)
'    ActualSecond = ActualSecond + 1
'    If ActualSecond = LastSecond Then End
'    LastSecond = ActualSecond
'End Sub

Private Sub TirarItem()
    If (ItemElegido > 0 And ItemElegido < MAX_INVENTORY_SLOTS + 1) Or (ItemElegido = FLAGORO) Then
        If UserInventory(ItemElegido).Amount = 1 Then
            SendData "TI" & ItemElegido & "," & 1
        Else
            If UserInventory(ItemElegido).Amount > 1 Then
                frmCantidad.Show
            End If
        End If
    End If


End Sub

Private Sub AgarrarItem()
    SendData "AG"

End Sub

Private Sub UsarItem()
    If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then
        SendData "USA" & ItemElegido & " " & RandomNumber(1, 5)    '< la data random que recibe
    End If

End Sub
Public Sub EquiparItem()

    If (ItemElegido > 0) And (ItemElegido < MAX_INVENTORY_SLOTS + 1) Then _
       SendData "EQUI" & ItemElegido

End Sub
Private Sub lblLanzar_Click()
    Call Sound.Sound_Play(SND_CLICK)

    If lstHechizos.List(lstHechizos.ListIndex) <> "Nada" And TiempoTranscurrido(LastHechizo) >= IntervaloSpell And TiempoTranscurrido(Hechi) >= IntervaloSpell / 4 Then
        Call SendData("LH" & lstHechizos.ListIndex + 1 & " " & RandomNumber(1, 5))
        Call SendData("UK" & Magia)
    End If

End Sub
Private Sub lblInfo_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call SendData("INFS" & lstHechizos.ListIndex + 1)
End Sub
Private Sub renderer_DblClick()
    If Not frmForo.Visible Then
        SendData "RC" & tX & "," & tY
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If (Not SendTxt.Visible) Then
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then

            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    If Opciones.sMusica = 1 Then
                        Opciones.sMusica = 0
                        Sound.Music_Stop
                        Call WriteVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG AUDIO", "Musica", 0)
                        Exit Sub
                    Else
                        Opciones.sMusica = 1
                        Sound.Music_Play
                        Call WriteVar(App.Path & "\RECURSOS\Init\Opciones.opc", "CONFIG AUDIO", "Musica", 1)
                        Exit Sub
                    End If    'X
                    '  Call WriteVar(App.Path & "\RECURSOS/Init/Opciones.opc", "CONFIG", "Musica", Str(Musica))
                    '/Mp3

                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem    'X

                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem    'X

                Case vbKeyK:
                    Call SendData("/SEGUROCLAN")

                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    Call SendData("UK" & Domar)    'X

                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    Call SendData("UK" & Robar)    'X

                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    Call SendData("UK" & Ocultarse)    'X

                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem    'X

                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    If Not NoPuedeUsar Then
                        NoPuedeUsar = True
                        Call UsarItem
                    End If    'X

                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    Call SendData("RPU")
                    '..........................ShaFTeR..........................
                Case CustomKeys.BindedKey(eKeyType.mKeyNormal)
                    frmMain.modo = "1 Normal"
                    If EligiendoWhispereo Then
                        EligiendoWhispereo = False
                        MousePointer = 1
                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeySusurrar)
                    Call AddtoRichTextBox(frmMain.rectxt, "Has click sobre el usuario al que quieres susurrar.", 255, 255, 255, 1, 0)
                    frmMain.modo = "2 Susurrar"
                    MousePointer = 2
                    EligiendoWhispereo = True

                Case CustomKeys.BindedKey(eKeyType.mKeyClan)
                    frmMain.modo = "3 Clan"
                    If EligiendoWhispereo Then
                        EligiendoWhispereo = False
                        MousePointer = 1
                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeyGrito)
                    frmMain.modo = "4 Grito"
                    If EligiendoWhispereo Then
                        EligiendoWhispereo = False
                        MousePointer = 1
                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeyRol)
                    frmMain.modo = "5 Rol"
                    If EligiendoWhispereo Then
                        EligiendoWhispereo = False
                        MousePointer = 1
                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeyParti)
                    frmMain.modo = "6 Party"
                    If EligiendoWhispereo Then
                        EligiendoWhispereo = False
                        MousePointer = 1
                    End If
                Case CustomKeys.BindedKey(eKeyType.mkeyRmsg)
                    frmMain.modo = "7 RMSG"
                    If EligiendoWhispereo Then
                        EligiendoWhispereo = False
                        MousePointer = 1
                    End If

                Case CustomKeys.BindedKey(eKeyType.mKeyGlobal)
                    frmMain.modo = "8 Global"
                    If EligiendoWhispereo Then
                        EligiendoWhispereo = False
                        MousePointer = 1
                    End If

                    '..........................ShaFTeR..........................

                    '          Case CustomKeys.BindedKey(eKeyType.mKeyToggleSafeMode)
                Case CustomKeys.BindedKey(eKeyType.mKeyParty)
                    frmParty.ListaIntegrantes.Clear
                    LlegoParty = False
                    Call SendData("PARINF")
                    Do While Not LlegoParty
                        DoEvents
                    Loop
                    frmParty.Visible = True
                    frmParty.SetFocus
                    LlegoParty = False
                    '   Case CustomKeys.BindedKey(eKeyType.mKeyTalkWithGuild)
                Case CustomKeys.BindedKey(eKeyType.mKeyInvi)
                    Call SendData("/INVISIBLE")
                    '   Case CustomKeys.BindedKey(eKeyType.mKeyToggleFPS)
                Case CustomKeys.BindedKey(eKeyType.mKeyTakeScreenShot)
                    'TomarFoto

                Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
                    Call frmOpciones.Show(vbModeless, frmMain)

                Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
                    Call SendData("/MEDITAR")    'X

                    '   Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)


                Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
                    Call SendData("#B")    'X

                Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
                    If (TiempoTranscurrido(LastGolpe) >= IntervaloGolpe) And (TiempoTranscurrido(Golpeo) >= IntervaloGolpe / 4) And (Not UserDescansar) And _
                       (Not UserMeditar) Then
                        Call SendData("AT")
                        Golpeo = Timer
                    End If    'X

                Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
                    If Not frmCantidad.Visible Then
                        SendTxt.Visible = True
                        SendTxt.SetFocus
                    End If    'X

                    'Standelf
                Case CustomKeys.BindedKey(eKeyType.mKeyUnlock)
                    Call SendData("(A")    'X
            End Select
        End If
    End If
End Sub

Sub Form_Load()
    If App.exeName = "Lhirius AO" Then Call StartURLDetect(rectxt.hwnd, Me.hwnd)
    PuertoDelServidor = 12345

    frmMain.Caption = "Lhirius AO" & " V " & App.Major & "." & App.Minor

    Me.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\principal.jpg")
    frmMain.imgFondoInvent.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\inventario.jpg")
    frmMain.imgFondoHechizos.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Hechizos.jpg")
    frmMain.STAShp.Picture = LoadPicture(App.Path & "\RECURSOS\Graficos\Grh\energia.jpg")
    frmMain.Hpshp.Picture = LoadPicture(App.Path & "\RECURSOS\Graficos\Grh\salud.jpg")
    frmMain.MANShp.Picture = LoadPicture(App.Path & "\RECURSOS\Graficos\Grh\mana.jpg")
    frmMain.AGUAsp.Picture = LoadPicture(App.Path & "\RECURSOS\Graficos\Grh\sed.jpg")
    frmMain.COMIDAsp.Picture = LoadPicture(App.Path & "\RECURSOS\Graficos\Grh\hambre.jpg")

End Sub
Private Sub lstHechizos_KeyDown(KeyCode As Integer, Shift As Integer)

    KeyCode = 0

End Sub
Private Sub lstHechizos_KeyPress(KeyAscii As Integer)

    KeyAscii = 0

End Sub
Private Sub lstHechizos_KeyUp(KeyCode As Integer, Shift As Integer)

    KeyCode = 0

End Sub
Private Sub Image1_Click(Index As Integer)
    Call Sound.Sound_Play(SND_CLICK)

    Select Case Index
        Case 0
            Call frmOpciones.Show(vbModeless, frmMain)
        Case 1
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            LlegoMinist = False
            SendData "ATRI"
            SendData "ESKI"
            SendData "FAMA"
            Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama Or Not LlegoMinist
                DoEvents
            Loop
            frmEstadisticas.Iniciar_Labels
            frmEstadisticas.Show
            LlegaronAtrib = False
            LlegaronSkills = False
            LlegoFama = False
            LlegoMinist = False
        Case 2
            If frmGuildLeader.Visible Then frmGuildLeader.Visible = False
            If frmGuildsNuevo.Visible Then frmGuildsNuevo.Visible = False
            If frmGuildAdm.Visible Then frmGuildAdm.Visible = False
            Call SendData("GLINFO")
        Case 3
            frmMapa.Visible = True
    End Select

End Sub

Private Sub Image3_Click()

    Call Sound.Sound_Play(SND_CLICK)
    If MsgBox("¿Desea salir de Lhirius AO?", vbYesNo, "Salir de Lhirius AO") = vbYes Then
        Call Sound.Sound_Play(SND_CLICK)
        Call SendData("#B")
    Else
        Call Sound.Sound_Play(SND_CLICK)
    End If

End Sub

Private Sub Label1_Click()
    LlegaronSkills = False
    SendData "ESKI"

    Do While Not LlegaronSkills
        DoEvents
    Loop

    Dim i As Integer
    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i
    Alocados = SkillPoints
    frmSkills3.Puntos.Caption = SkillPoints
    frmSkills3.Show
End Sub

Private Sub RecTxt_Change()
    On Error Resume Next

    If SendTxt.Visible Then
        SendTxt.SetFocus
    ElseIf (Not frmComerciar.Visible) And _
           (Not frmSkills3.Visible) And _
           (Not frmForo.Visible) And _
           (Not frmEstadisticas.Visible) And _
           (Not frmCantidad.Visible) Then
        ' Picture1.SetFocus
    End If

End Sub
Private Sub SendTxt_Change()

    stxtbuffer = SendTxt.Text

End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then
        KeyAscii = 0
        'frmMain.SetFocus
    End If
End Sub

Private Sub Socket1_Connect()

'Second.Enabled = True

'Call SendData(gsEnviarID)
    If EstadoLogin = CrearNuevoPj Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = CrearAccount Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = Dados Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = LoginAccount Then
        Call SendData("gIvEmEvAlcOde")
    ElseIf EstadoLogin = Normal Then
        Call SendData("gIvEmEvAlcOde")
    End If

End Sub
Private Sub Socket1_Disconnect()
    On Error Resume Next
    Connected = False

    Socket1.Cleanup

    frmConnect.MousePointer = vbNormal
    frmCrearPersonaje.Visible = False
    frmConnect.Visible = True

    frmMain.Visible = False

    Pausa = False
    UserMeditar = False

    UserSexo = ""
    UserRaza = ""
    UserEmail = ""
    bO = 100

    Dim i As Integer
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i

    SkillPoints = 0
    Alocados = 0

    Dialogos.UltimoDialogo = 0
    Dialogos.CantidadDialogos = 0
    Sound.Sound_Stop_All
    Sound.Ambient_Stop
    If Opciones.sMusica <> CONST_DESHABILITADA Then
        If Opciones.sMusica <> CONST_DESHABILITADA Then
            Sound.Fading = 350
            Sound.Music_Load (1)
            'Sound.Sound_Render
            Sound.Music_Play
        End If
    End If
    YaLoguio = False
End Sub
Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)

    Select Case ErrorCode
        Case 24036
            Call MsgBox("Por favor espere, intentando completar conexión.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
            Exit Sub

        Case 24053
            Call MsgBox("Conexion Perdida.", vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")

        Case 24060
            Call MsgBox("Tiempo de Espera Terminado.", vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")

        Case Else
            Call MsgBox(ErrorString, vbApplicationModal + vbCritical + vbOKOnly + vbDefaultButton1, "Error")

    End Select

    frmConnect.MousePointer = 1
    Response = 0
    'LastSecond = 0
    'Second.Enabled = False

    frmMain.Socket1.Disconnect

    If Not frmCrearPersonaje.Visible Then
        frmConnect.Show
    Else
        frmCrearPersonaje.MousePointer = 0
    End If

End Sub
Private Sub Socket1_Read(DataLength As Integer, IsUrgent As Integer)
    Dim loopc As Integer

    Dim RD As String
    Dim rBuffer(1 To 500) As String

    Static TempString As String

    Dim CR As Integer
    Dim tChar As String
    Dim sChar As Integer

    Call Socket1.Read(RD, DataLength)

    If TempString <> "" Then
        RD = TempString & RD
        TempString = ""
    End If

    sChar = 1

    For loopc = 1 To Len(RD)
        tChar = mid$(RD, loopc, 1)

        If tChar = ENDC Then
            CR = CR + 1
            rBuffer(CR) = mid$(RD, sChar, loopc - sChar)
            sChar = loopc + 1
        End If

    Next loopc

    If Len(RD) - (sChar - 1) <> 0 Then TempString = mid$(RD, sChar, Len(RD))

    For loopc = 1 To CR
        Call HandleData(rBuffer(loopc))
    Next loopc

End Sub
Private Sub SOS_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SOS.ForeColor = &HFFFF&
    SOS.BackColor = &HFF00&
    Call SendData("/DAMESOS")
End Sub

Private Sub SOS_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SOS.ForeColor = &HFFFF&
    SOS.BackColor = &H4000&
End Sub

Private Sub Timerasd_Timer()
    Call UsarItem
End Sub

Private Sub TimerPoteo_Timer()
    PuedePotear = True
End Sub

Private Sub TimerPoteoClick_Timer()
    PuedePotearClick = True
End Sub
Private Sub RecTxt_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Particula Then
        Particula = False
        effect(EIndex).Used = False
        effect(EIndex2).Used = False
        EIndex = 0
        EIndex2 = 0
    End If
End Sub
Private Sub renderer_Click()

    If Cartel Then Cartel = False

    If Comerciando = 0 Then
        Call ConvertCPtoTP(MouseX, MouseY, tX, tY)
        If Abs(UserPos.Y - tY) > 6 Then Exit Sub
        If Abs(UserPos.X - tX) > 8 Then Exit Sub
        If EligiendoWhispereo Then
            Call SendData("WH" & tX & "," & tY)
            EligiendoWhispereo = False
            Exit Sub
        End If

        If UsingSkill = 0 Then
            SendData "LC" & tX & "," & tY
        Else
            frmMain.MousePointer = vbDefault
            If UsingSkill = Magia Then

                frmMain.MousePointer = vbDefault
                If Particula Then
                    Particula = False
                    effect(EIndex).Used = False
                    effect(EIndex2).Used = False
                    EIndex = 0
                    EIndex2 = 0
                End If

                If (TiempoTranscurrido(LastHechizo) < IntervaloSpell Or TiempoTranscurrido(Hechi) < IntervaloSpell / 4) Then
                    Exit Sub
                Else: Hechi = Timer
                End If
            ElseIf UsingSkill = Proyectiles Then
                If (TiempoTranscurrido(LastFlecha) < IntervaloFlecha Or TiempoTranscurrido(Flecho) < IntervaloFlecha / 4) Then
                    Exit Sub
                Else: Flecho = Timer
                End If
            End If
            Call SendData("WLC" & tX & "," & tY & "," & UsingSkill)
            UsingSkill = 0
        End If
    End If

    If boton = vbRightButton Then
        Call SendData("/TELEPLOC")
        DibujarPuntoMinimap
    End If
    boton = 0

End Sub
