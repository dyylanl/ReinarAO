VERSION 5.00
Begin VB.Form frmUserRequest 
   BackColor       =   &H00111720&
   BorderStyle     =   0  'None
   Caption         =   "Peticion"
   ClientHeight    =   2895
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   4755
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   4755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H80000009&
      Height          =   1575
      Left            =   600
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   3615
   End
   Begin VB.Image command1 
      Height          =   375
      Left            =   4440
      MouseIcon       =   "frmUserRequest.frx":0000
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   375
   End
End
Attribute VB_Name = "frmUserRequest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Unload Me
End Sub

Public Sub recievePeticion(ByVal P As String)

    Text1 = Replace(P, "�", vbCrLf)
    Me.Show vbModeless, frmMain

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
Private Sub Form_Load()

    Me.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\peticion.gif")

End Sub

