VERSION 5.00
Begin VB.Form frmConnect 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Lhirius AO"
   ClientHeight    =   8985
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11985
   ClipControls    =   0   'False
   FillColor       =   &H00000040&
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmConnect.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   599
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   799
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4080
      MaskColor       =   &H00004080&
      TabIndex        =   1
      Top             =   5610
      UseMaskColor    =   -1  'True
      Width           =   158
   End
   Begin VB.Timer Timer2 
      Left            =   2520
      Top             =   1320
   End
   Begin VB.Timer Timer1 
      Left            =   2040
      Top             =   1320
   End
   Begin VB.PictureBox RenderConnect 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   9000
      Left            =   0
      ScaleHeight     =   600
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   800
      TabIndex        =   0
      Top             =   0
      Width           =   12000
      Begin VB.Frame Frame3 
         Caption         =   "Salir ----->"
         Height          =   375
         Left            =   4080
         TabIndex        =   4
         Top             =   7560
         Width           =   1095
      End
      Begin VB.Frame Frame2 
         Caption         =   "Conectarse     ---->"
         Height          =   375
         Left            =   3840
         TabIndex        =   3
         Top             =   6120
         Width           =   1575
      End
      Begin VB.Frame Frame1 
         Caption         =   "Crear Cuenta ---->"
         Height          =   375
         Left            =   3960
         TabIndex        =   2
         Top             =   6840
         Width           =   1455
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   4920
         Top             =   7560
         Width           =   2295
      End
      Begin VB.Image imgWeb 
         Height          =   1695
         Left            =   2520
         MouseIcon       =   "frmConnect.frx":000C
         MousePointer    =   99  'Custom
         Top             =   240
         Width           =   7335
      End
      Begin VB.Image Image1 
         Height          =   375
         Index           =   1
         Left            =   3960
         MouseIcon       =   "frmConnect.frx":0316
         MousePointer    =   99  'Custom
         Top             =   6120
         Width           =   4050
      End
      Begin VB.Image Image1 
         Height          =   315
         Index           =   0
         Left            =   3960
         MouseIcon       =   "frmConnect.frx":0620
         MousePointer    =   99  'Custom
         Top             =   6840
         Width           =   4050
      End
      Begin VB.Image imgDato 
         Height          =   165
         Index           =   1
         Left            =   4590
         Top             =   5055
         Width           =   2895
      End
      Begin VB.Image imgDato 
         Height          =   165
         Index           =   0
         Left            =   4590
         Top             =   4575
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private ret As Long
Private LocalIP As Boolean
Private cont2 As Integer


Private Sub Check1_Click()
    If Check1.value = True Then
        Call WriteVar(App.Path & "\RECURSOS\INIT\Opciones.opc", "RECORDAR", "Nombre", "")
        Call WriteVar(App.Path & "\RECURSOS\INIT\Opciones.opc", "RECORDAR", "Password", "")
        Call WriteVar(App.Path & "\RECURSOS\INIT\Opciones.opc", "RECORDAR", "Check", "0")
        txtNombre = ""
        txtPasswd = ""
        MsgBox "Su cuenta no está guardada."
    End If
End Sub

Private Sub Command1_Click()
    Call Image2_Click
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

'    If KeyAscii = vbKeyReturn Then
'        Call Sound.Sound_Play(SND_CLICK)'

'        nombrecuent = txtNombre
'        If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect'

'        If frmConnect.MousePointer = 11 Then
'            frmConnect.MousePointer = 1
'            Exit Sub
'        End If



 '       UserName = txtNombre
 '       Dim aux As String
 '       aux = txtPasswd
 '       UserPassword = MD5String(aux)
 '       If CheckUserData(False) = True Then
 '           frmMain.Socket1.HostName = IPdelServidor
 '           frmMain.Socket1.RemotePort = PuertoDelServidor'

'            EstadoLogin = LoginAccount
'            Me.MousePointer = 11
'            frmMain.Socket1.Connect
'        End If
'    End If

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF1 Then
        If LocalIP = True Then    'es local la ip
            IPdelServidor = "186.60.181.156"    'Host
            MsgBox "IP Publica configurada."
            LocalIP = False
            Exit Sub
        Else
            IPdelServidor = "127.0.0.1"
            MsgBox "IP Local Configurada."
            LocalIP = True
            Exit Sub
        End If
    End If
    If KeyCode = 27 Then
        DeInitTileEngine
        End
    End If

End Sub


Private Sub Form_Load()
    On Error Resume Next
    LocalIP = False

    'Me.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\conectar.bmp")
    cont2 = 255
    ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED
    SetWindowLong Me.hwnd, GWL_EXSTYLE, ret
    Timer1.Interval = 1
    Timer2.Interval = 1
    Timer2.Enabled = False
    Timer1.Enabled = True
    EngineRun = False


    Dim j
    For Each j In Image1()
        j.Tag = "0"
    Next

    IntervaloPaso = 0.19

    If GetVar(App.Path & "\RECURSOS\INIT\Opciones.opc", "RECORDAR", "Check") = 1 Then
        Check1.value = 1
        txtNombre = GetVar(App.Path & "\RECURSOS\INIT\Opciones.opc", "RECORDAR", "Nombre")
        txtPasswd = GetVar(App.Path & "\RECURSOS\INIT\Opciones.opc", "RECORDAR", "Password")
        txtPasswd = DesEncriptar(txtPasswd)    'Gracias a Tonchitoz por esto
    Else
        Check1.value = 0
    End If
End Sub

Private Sub imgDato_Click(Index As Integer)

If Index = 0 Then FocoPasswd = False Else FocoPasswd = True

End Sub
Private Sub Image1_Click(Index As Integer)

    Call Sound.Sound_Play(SND_CLICK)

    Select Case Index
        Case 0

            EstadoLogin = CrearAccount
            frmMain.Socket1.HostName = IPdelServidor
            frmMain.Socket1.RemotePort = PuertoDelServidor
            Me.MousePointer = 11
            frmMain.Socket1.Connect
            If Opciones.sMusica <> CONST_DESHABILITADA Then
                If Opciones.sMusica <> CONST_DESHABILITADA Then
                    Sound.NextMusic = MUS_CrearPersonaje
                    Sound.Fading = 500
                End If
            End If
            
        Case 1
            If Check1.value = 1 Then
                Call WriteVar(App.Path & "\RECURSOS\INIT\Opciones.opc", "RECORDAR", "Nombre", txtNombre)
                Call WriteVar(App.Path & "\RECURSOS\INIT\Opciones.opc", "RECORDAR", "Password", Encriptar(txtPasswd))    'Gracias a TonchitoZ por esto.
                Call WriteVar(App.Path & "\RECURSOS\INIT\Opciones.opc", "RECORDAR", "Check", "1")
            End If
            nombrecuent = txtNombre
            If frmMain.Socket1.Connected Then frmMain.Socket1.Disconnect

            If frmConnect.MousePointer = 11 Then
                frmConnect.MousePointer = 1
                Exit Sub
            End If



            UserName = txtNombre
            Dim aux As String
            aux = txtPasswd
            UserPassword = MD5String(aux)
            If CheckUserData(False) = True Then
                frmMain.Socket1.HostName = IPdelServidor
                frmMain.Socket1.RemotePort = PuertoDelServidor

                EstadoLogin = LoginAccount
                Me.MousePointer = 11
                frmMain.Socket1.Connect
            End If

    End Select

End Sub

Private Sub Image2_Click()
    DeInitTileEngine
    End
End Sub

Private Sub imgWeb_Click()
    Call Sound.Sound_Play(SND_CLICK)
    Call ShellExecute(Me.hwnd, "open", "http://lhirius.zxq.net/test/index.php", "", "", 1)

End Sub

Private Sub RenderConnect_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then Call Image1_Click(1): Exit Sub

If FocoPasswd = True Then
 
 If KeyAscii = vbKeyBack Then
  If Len(txtPasswd) > 0 Then
   txtPasswd = Left(txtPasswd, Len(txtPasswd) - 1)
   txtPasswdAsteriscos = Left(txtPasswdAsteriscos, Len(txtPasswdAsteriscos) - 1)
   Exit Sub
  End If
  If Len(txtPasswdAsteriscos) = 0 Then Beep: Exit Sub
 End If
If Len(txtPasswd) < 30 Then txtPasswd = txtPasswd & Chr(KeyAscii) Else Beep
If Len(txtPasswdAsteriscos) < 30 Then txtPasswdAsteriscos = txtPasswdAsteriscos + "*"

ElseIf FocoPasswd = False Then
 
 If KeyAscii = vbKeyBack Then
  If Len(txtNombre) > 0 Then
   txtNombre = Left(txtNombre, Len(txtNombre) - 1)
   Exit Sub
  End If
  If Len(txtNombre) = 0 Then Beep: Exit Sub
 End If
If Len(txtNombre) < 30 Then txtNombre = txtNombre & Chr(KeyAscii) Else Beep

End If
End Sub

Private Sub RenderConnect_LostFocus()
If frmConnect.Visible Then
RenderConnect.SetFocus

If FocoPasswd Then
 imgDato_Click (0)
Else
 imgDato_Click (1)
End If
End If
End Sub

Private Sub Timer1_Timer()
    Static Cont As Integer
    Cont = Cont + 5
    If Cont > 255 Then
        Cont = 0
        Timer1.Enabled = False
    Else
        SetLayeredWindowAttributes Me.hwnd, 0, Cont, LWA_ALPHA
    End If
End Sub

Private Sub Timer2_Timer()
    cont2 = cont2 - 5
    If cont2 < 0 Then
        Timer2.Enabled = False
        End
    Else
        SetLayeredWindowAttributes Me.hwnd, 0, cont2, LWA_ALPHA
    End If
End Sub
