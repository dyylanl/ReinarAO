VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form FrmIntro 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7065
   ClientLeft      =   4170
   ClientTop       =   2565
   ClientWidth     =   12000
   Icon            =   "FrmIntro.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7065
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer2 
      Left            =   1920
      Top             =   360
   End
   Begin VB.Timer Timer1 
      Left            =   1320
      Top             =   360
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      CausesValidation=   0   'False
      DragMode        =   1  'Automatic
      Height          =   5115
      Left            =   4800
      TabIndex        =   2
      Top             =   1200
      Width           =   6735
      ExtentX         =   11880
      ExtentY         =   9022
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
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   6720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ONLINE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   360
      Left            =   1560
      TabIndex        =   1
      Top             =   6600
      Width           =   1155
   End
   Begin VB.Image Image1 
      Height          =   975
      Left            =   720
      Top             =   1200
      Width           =   3735
   End
   Begin VB.Image Image2 
      Height          =   855
      Left            =   840
      MouseIcon       =   "FrmIntro.frx":324A
      MousePointer    =   99  'Custom
      Top             =   2640
      Width           =   3495
   End
   Begin VB.Image Image6 
      Height          =   615
      Left            =   11400
      MouseIcon       =   "FrmIntro.frx":3554
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   615
   End
   Begin VB.Image Image5 
      Height          =   975
      Left            =   600
      MouseIcon       =   "FrmIntro.frx":385E
      MousePointer    =   99  'Custom
      Top             =   5280
      Width           =   3855
   End
   Begin VB.Image Image4 
      Height          =   975
      Left            =   720
      MouseIcon       =   "FrmIntro.frx":3B68
      MousePointer    =   99  'Custom
      Top             =   3960
      Width           =   3735
   End
End
Attribute VB_Name = "FrmIntro"
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

Private Sub Form_Load()
FrmMusica.Show
Call WebBrowser1.Navigate("http://genius-ao.forosactivos.com/Noticias-h2.htm")
Me.Picture = LoadPicture(App.Path & "\Interfaces\MenuPrincipal.gif")
Dim corriendo As Integer
Dim i As Long
Dim proc As PROCESSENTRY32
Dim snap As Long
Dim pepe As String
   cont2 = 255
    ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, ret
    Timer1.Interval = 1
    Timer2.Interval = 1
    Timer2.Enabled = False
    Timer1.Enabled = True
Dim exeName As String
snap = CreateToolhelpSnapshot(TH32CS_SNAPALL, 0)
proc.dwSize = Len(proc)
theloop = ProcessFirst(snap, proc)
i = 0
While theloop <> 0
    exeName = proc.szexeFile
    Text1.Text = proc.szexeFile
    If Text1.Text = "DragoonAONoDinamico.exe" Or Text1.Text = "DragoonAO.exe" Then
        corriendo = corriendo + 1
        Text1.Text = ""
    End If
    i = i + 1
    theloop = ProcessNext(snap, proc)
Wend
CloseHandle snap

End Sub

Private Sub Image1_Click()
'Cerramos formulario
 
Me.Hide
 
'Petin _
Iniciamos Juego
Call Main
End Sub

Private Sub Image2_Click()
ShellExecute Me.hWnd, "open", App.Path & "/Update.exe", "", "", 1
Unload Me
End Sub

Private Sub Image3_Click()
ShellExecute Me.hWnd, "open", App.Path & "/aosetup.exe", "", "", 1
End Sub

Private Sub Image4_Click()
ShellExecute Me.hWnd, "open", "http://www.lhirius-ao.com.ar", "", "", 1

End Sub

Private Sub Image5_Click()
ShellExecute Me.hWnd, "open", "http://www.genius-ao.forosactivos.com", "", "", 1

End Sub

Private Sub Image6_Click()
Cancel = True
Timer2.Enabled = True
End

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving = False And Button = vbLeftButton Then

      Dx3 = X

      dy = Y

      bmoving = True

   End If

   

End Sub

 

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If bmoving And ((X <> Dx3) Or (Y <> dy)) Then

      Move Left + (X - Dx3), Top + (Y - dy)

   End If
   End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

   If Button = vbLeftButton Then

      bmoving = False

   End If

   

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

