VERSION 5.00
Begin VB.Form frmMSG 
   BorderStyle     =   0  'None
   Caption         =   "GM Messenger"
   ClientHeight    =   7230
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   6450
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   6450
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox mensaje 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   1575
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4680
      Width           =   5175
   End
   Begin VB.TextBox GM 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   330
      Left            =   2280
      TabIndex        =   2
      Text            =   "Cualquier GM disponible"
      Top             =   2640
      Width           =   2775
   End
   Begin VB.ComboBox categoria 
      Appearance      =   0  'Flat
      BackColor       =   &H00111720&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   345
      ItemData        =   "frmMSG.frx":0000
      Left            =   3360
      List            =   "frmMSG.frx":0013
      TabIndex        =   1
      Top             =   3720
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   3480
      MouseIcon       =   "frmMSG.frx":0065
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   2295
   End
   Begin VB.Image command1 
      Height          =   375
      Left            =   840
      MouseIcon       =   "frmMSG.frx":036F
      MousePointer    =   99  'Custom
      Top             =   6480
      Width           =   1935
   End
End
Attribute VB_Name = "frmMSG"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'FénixAO en DX8 by ·Parra, Thusing y DarkTester

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()

Me.picture = LoadPicture(App.Path & "\Interfaces\SGM.gif")

End Sub
Private Sub Image1_Click()
Dim GMs As String

If categoria.ListIndex = -1 Then
    MsgBox "El motivo del mensaje no es válido"
    Exit Sub
End If

If Len(mensaje.Text) > 250 Then
    MsgBox "La longitud del mensaje debe tener menos de 250 carácteres."
    Exit Sub
End If

If Len(GM.Text) = 0 Or GM.Text = "Cualquier GM disponible" Then
    GMs = "Ninguno"
Else: GMs = GM.Text
End If

If Len(mensaje.Text) = 0 Then
    MsgBox "Debes ingresar un mensaje."
    Exit Sub
End If

Call SendData("GM" & GMs & "¬" & categoria.List(categoria.ListIndex) & "¬" & mensaje.Text)

If NoMandoElMsg = 0 Then
    mensaje.Text = ""
    GM.Text = "Cualquier GM disponible"
    categoria.List(categoria.ListIndex) = ""
    AddtoRichTextBox frmMain.rectxt, "El mensaje fue enviado. Dentro de algunas horas recibirás la respuesta. Rogamos tengas paciencia y no escribas más de un mensaje sobre el mismo tema.", 252, 151, 53, 1, 0
    Unload Me
Else
    Call MsgBox("El mensaje es demasiado largo, por favor resumilo.")
End If

End Sub


Private Sub mensaje_Change()
mensaje.Text = LTrim(mensaje.Text)
End Sub


Private Sub mensaje_KeyPress(KeyAscii As Integer)

If (KeyAscii <> 209) And (KeyAscii <> 241) And (KeyAscii <> 8) And (KeyAscii <> 32) And (KeyAscii <> 164) And (KeyAscii <> 165) Then
    If (index <> 6) And ((KeyAscii < 40 Or KeyAscii > 122) Or (KeyAscii > 90 And KeyAscii < 96)) Then
        KeyAscii = 0
    End If
End If

 KeyAscii = Asc((Chr(KeyAscii)))
End Sub
