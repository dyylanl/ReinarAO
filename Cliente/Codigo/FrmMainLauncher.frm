VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form FrmMainLauncher 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Lhirius AO"
   ClientHeight    =   8490
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7110
   Icon            =   "FrmMainLauncher.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   7110
   StartUpPosition =   1  'CenterOwner
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   6015
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   6615
      ExtentX         =   11668
      ExtentY         =   10610
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      ForeColor       =   &H8000000E&
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   6450
      Width           =   6615
   End
   Begin VB.Image ImgSalir 
      Height          =   615
      Left            =   3600
      Top             =   7680
      Width           =   3210
   End
   Begin VB.Image ImgJugar 
      Height          =   570
      Left            =   240
      Top             =   6840
      Width           =   6570
   End
   Begin VB.Image imgForo 
      Height          =   615
      Left            =   240
      Top             =   7680
      Width           =   3240
   End
End
Attribute VB_Name = "FrmMainLauncher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const Directorio As String = "\RECURSOS\Graficos\Launcher\"

Dim f As Integer

Private Sub Form_Load()

    WebBrowser1.Navigate "http://lhirius-ao.forosactivos.com/h2-noticias"
    ImgSalir.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Salir_N.jpg")
    ImgJugar.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Jugar_N.jpg")
    imgForo.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Foro_N.jpg")
    Me.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main.jpg")

    'Call Analizar
    IPdelServidor = "186.60.181.156"    'Host
End Sub

Private Sub Analizar()
    On Error Resume Next
    Dim ix As Integer, tX As Integer, DifX As Integer

    'lEstado.Caption = "Obteniendo datos..."
    lblEstado.Caption = "Buscando Actualizaciones..."
    lblEstado.ForeColor = vbGreen

    ix = Inet1.OpenURL("http://lhirius-ao.gzpot.com/Autoupdate/VEREXE.txt")    'Host
    tX = LeerInt(App.Path & "\RECURSOS\INIT\Update.ini")


    DifX = ix - tX

    If Not (DifX = 0) Then
        lblEstado.Caption = "Hay " & DifX & " actualizaciones disponibles."
        lblEstado.ForeColor = vbRed
    Else
        lblEstado.Caption = "Lhirius AO está actualizado. Pulsa el botón Jugar."
        lblEstado.ForeColor = vbGreen
    End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    ImgSalir.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Salir_N.jpg")
    ImgJugar.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Jugar_N.jpg")
    imgForo.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Foro_N.jpg")

End Sub

Private Sub imgForo_Click()
    Call ShellExecute(Me.hwnd, "Open", "http://www.lhirius-ao.forosactivos.com/", &O0, &O0, 1)
End Sub

Private Sub imgForo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgForo.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Foro_A.jpg")
End Sub

Private Sub imgForo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgForo.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Foro_I.jpg")
End Sub

Private Sub ImgJugar_Click()

    If lblEstado.ForeColor = vbRed Then
        Call MsgBox("Se abrirá el AutoUpdate para actualizar Lhirius AO a la versión mas actual.", vbInformation, "Atención")
        Unload FrmMainLauncher
        Call ShellExecute(Me.hwnd, "open", App.Path & "/Autoupdate.exe", "", "", 1)
        End
        Exit Sub
    Else
        Unload FrmMainLauncher
        Call Main
        Exit Sub
    End If

End Sub

Private Sub ImgJugar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgJugar.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Jugar_A.jpg")
End Sub

Private Sub ImgJugar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgJugar.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Jugar_I.jpg")
End Sub

Private Sub ImgSalir_Click()

    End

End Sub

Private Sub ImgSalir_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgSalir.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Salir_A.jpg")
End Sub

Private Sub ImgSalir_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ImgSalir.Picture = LoadPicture(App.Path & Directorio & "Launcher_Main_Salir_I.jpg")
End Sub

Private Function LeerInt(ByVal Ruta As String) As Integer
    f = FreeFile
    Open Ruta For Input As f
    LeerInt = Input$(LOF(f), #f)
    Close #f
End Function

