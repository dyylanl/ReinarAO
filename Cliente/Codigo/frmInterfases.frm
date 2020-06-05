VERSION 5.00
Begin VB.Form frmInterfases 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Elejir Interfaces en Lhirius AO"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   6165
   Icon            =   "frmInterfases.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Left            =   5640
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Left            =   5160
      Top             =   360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   4800
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aplicar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      MaskColor       =   &H00808080&
      TabIndex        =   5
      Top             =   4800
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "Interfases Disponibles"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   720
      Width           =   2655
      Begin VB.OptionButton Option7 
         BackColor       =   &H00000000&
         Caption         =   "Muerte del Alien"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   3000
         Width           =   1575
      End
      Begin VB.OptionButton Option6 
         BackColor       =   &H00000000&
         Caption         =   "Batalla Espacial"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   2520
         Width           =   1575
      End
      Begin VB.OptionButton Option5 
         BackColor       =   &H00000000&
         Caption         =   "Mision Complete"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         BackColor       =   &H00000000&
         Caption         =   "Cazadora  Night"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1560
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00000000&
         Caption         =   "Azul y Oro"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "The History"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   2295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Interface Clasica"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Vista Previa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   495
      Left            =   3600
      TabIndex        =   11
      Top             =   3720
      Width           =   2655
   End
   Begin VB.Image Image1 
      Height          =   2295
      Left            =   2880
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      Caption         =   "Cambiar Interfaces"
      BeginProperty Font 
         Name            =   "Tw Cen MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   420
      Left            =   1680
      TabIndex        =   7
      Top             =   120
      Width           =   2850
   End
End
Attribute VB_Name = "frmInterfases"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const LWA_COLORKEY = &H1
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private ret As Long
Private cont2 As Integer

Private Sub Command1_Click()
If Option1 = True Then
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\Principal.gif")
frmMain.imgFondoInvent.Picture = LoadPicture(App.Path & "\Interfaces\centronuevoinventario.gif")
frmMain.imgFondoHechizos.Picture = LoadPicture(App.Path & "\Interfaces\centronuevoHechizos.gif")
Call WriteVar(App.Path & "/Init/Interfases.ini", "ELEGIDA", "Interfase", 1)
End If

If Option2 = True Then
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\Principal1.gif")
frmMain.imgFondoInvent.Picture = LoadPicture(App.Path & "\Interfaces\centronuevoinventario1.gif")
frmMain.imgFondoHechizos.Picture = LoadPicture(App.Path & "\Interfaces\centronuevohechizos1.gif")
Call WriteVar(App.Path & "/Init/Interfases.ini", "ELEGIDA", "Interfase", 2)
End If

If Option3 = True Then
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\Principal2.gif")
frmMain.imgFondoInvent.Picture = LoadPicture(App.Path & "\Interfaces\centronuevoinventario2.gif")
frmMain.imgFondoHechizos.Picture = LoadPicture(App.Path & "\Interfaces\centronuevohechizos2.gif")
Call WriteVar(App.Path & "/Init/Interfases.ini", "ELEGIDA", "Interfase", 3)
End If

If Option4 = True Then
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\Principal3.gif")
frmMain.imgFondoInvent.Picture = LoadPicture(App.Path & "\Interfaces\centronuevoinventario3.gif")
frmMain.imgFondoHechizos.Picture = LoadPicture(App.Path & "\Interfaces\centronuevohechizos3.gif")
Call WriteVar(App.Path & "/Init/Interfases.ini", "ELEGIDA", "Interfase", 4)
End If

If Option5 = True Then
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\Principal4.gif")
frmMain.imgFondoInvent.Picture = LoadPicture(App.Path & "\Interfaces\centronuevoinventario4.gif")
frmMain.imgFondoHechizos.Picture = LoadPicture(App.Path & "\Interfaces\centronuevohechizos4.gif")
Call WriteVar(App.Path & "/Init/Interfases.ini", "ELEGIDA", "Interfase", 5)
End If

If Option6 = True Then
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\Principal5.gif")
frmMain.imgFondoInvent.Picture = LoadPicture(App.Path & "\Interfaces\centronuevoinventario5.gif")
frmMain.imgFondoHechizos.Picture = LoadPicture(App.Path & "\Interfaces\centronuevohechizos5.gif")
Call WriteVar(App.Path & "/Init/Interfases.ini", "ELEGIDA", "Interfase", 6)
End If

If Option7 = True Then
frmMain.Picture = LoadPicture(App.Path & "\Interfaces\Principal6.gif")
frmMain.imgFondoInvent.Picture = LoadPicture(App.Path & "\Interfaces\centronuevoinventario6.gif")
frmMain.imgFondoHechizos.Picture = LoadPicture(App.Path & "\Interfaces\centronuevohechizos6.gif")
Call WriteVar(App.Path & "/Init/Interfases.ini", "ELEGIDA", "Interfase", 7)
End If

End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
   cont2 = 255
    ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, ret
    Timer1.Interval = 1
    Timer2.Interval = 1
    Timer2.Enabled = False
    Timer1.Enabled = True
Image1.Stretch = True
End Sub

Private Sub Option1_Click()
Image1.Picture = LoadPicture(App.Path & "\Interfaces\Principal.gif")
End Sub

Private Sub Option2_Click()
Image1.Picture = LoadPicture(App.Path & "\Interfaces\Principal1.gif")
End Sub

Private Sub Option3_Click()
Image1.Picture = LoadPicture(App.Path & "\Interfaces\Principal2.gif")
End Sub

Private Sub Option4_Click()
Image1.Picture = LoadPicture(App.Path & "\Interfaces\Principal3.gif")
End Sub

Private Sub Option5_Click()
Image1.Picture = LoadPicture(App.Path & "\Interfaces\Principal4.gif")
End Sub

Private Sub Option6_Click()
Image1.Picture = LoadPicture(App.Path & "\Interfaces\Principal5.gif")
End Sub

Private Sub Option7_Click()
Image1.Picture = LoadPicture(App.Path & "\Interfaces\Principal6.gif")
End Sub
Private Sub Timer1_Timer()
    Static Cont As Integer
    Cont = Cont + 5
    If Cont > 255 Then
        Cont = 0
        Timer1.Enabled = False
    Else
        SetLayeredWindowAttributes Me.hWnd, 0, Cont, LWA_ALPHA
    End If
End Sub
 
Private Sub Timer2_Timer()
    cont2 = cont2 - 5
    If cont2 < 0 Then
        Timer2.Enabled = False
        End
    Else
        SetLayeredWindowAttributes Me.hWnd, 0, cont2, LWA_ALPHA
    End If
End Sub
