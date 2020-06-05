VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMP3 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Reproductor MP3"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4650
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmMP3.frx":0000
   ScaleHeight     =   1950
   ScaleWidth      =   4650
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Tu Musica en MP3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   4215
   End
   Begin VB.Image Image5 
      Height          =   615
      Left            =   600
      Top             =   960
      Width           =   615
   End
   Begin VB.Image Image4 
      Height          =   615
      Left            =   2040
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   2760
      Top             =   960
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   3360
      Top             =   960
      Width           =   615
   End
   Begin VB.Image Image1 
      Height          =   615
      Left            =   1320
      Top             =   960
      Width           =   615
   End
End
Attribute VB_Name = "frmMP3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Función Api GetShortPathName para obtener _
los paths de los archivos en formato corto
Private Declare Function GetShortPathName _
    Lib "kernel32" _
    Alias "GetShortPathNameA" ( _
        ByVal lpszLongPath As String, _
        ByVal lpszShortPath As String, _
        ByVal lBuffer As Long) As Long

'Función Api mciExecute para reproducir los archivos de música
Private Declare Function mciExecute _
    Lib "winmm.dll" ( _
        ByVal lpstrCommand As String) As Long
Dim ret As Long, Path As String
'Le pasamos el comando Close a MciExecute para cerrar el dispositivo
Private Sub Form_Unload(Cancel As Integer)
    mciExecute "Close All"
End Sub

'Sub que obtiene el path corto del archivo a reproducir
Private Sub PathCorto(Archivo As String)
Dim temp As String * 250 'Buffer
    Path = String(255, 0)
    'Obtenemos el Path corto
    ret = GetShortPathName(Archivo, temp, 164)
    'Sacamos los nulos al path
    Path = Replace(temp, Chr(0), "")
End Sub

'Procedimiento que ejecuta el comando con el Api mciExecute
'************************************************************
Private Sub ejecutar(comando As String)
    If Path = "" Then MsgBox "Error", vbCritical: Exit Sub
    'Llamamos a mciExecute pasandole un string que tiene el comando y la ruta

    mciExecute comando & Path

End Sub

Private Sub Image1_Click()
    ejecutar ("Pause ")
End Sub

Private Sub Image2_Click()
frmMP3.Hide
End Sub

Private Sub Image3_Click()
    With CommonDialog1
        .Filter = "Archivos Mp3|*.mp3|Archivos Wav|*.wav|Archivos MIDI|*.mid"
        .ShowOpen
        If .filename = "" Then
            Exit Sub
        Else
            'Le pasamos a la sub que obtiene con _
            el Api GetShortPathName el nombre corto del archivo
            PathCorto .filename
            Label1 = .filename
            'cerramos todo
            mciExecute "Close All"
            'Para Habilitar y deshabilitar botones
        End If
    End With
End Sub

Private Sub Image4_Click()
    ejecutar ("Stop ")
End Sub

Private Sub Image5_Click()
    ejecutar ("Play ")
End Sub

