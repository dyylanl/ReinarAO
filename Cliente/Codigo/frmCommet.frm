VERSION 5.00
Begin VB.Form frmCommet 
   BorderStyle     =   0  'None
   Caption         =   "Oferta de paz"
   ClientHeight    =   3825
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   5700
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
   ScaleHeight     =   3825
   ScaleWidth      =   5700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      ForeColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   700
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   950
      Width           =   4335
   End
   Begin VB.Image command2 
      Height          =   255
      Left            =   5400
      MouseIcon       =   "frmCommet.frx":0000
      MousePointer    =   99  'Custom
      Top             =   0
      Width           =   255
   End
   Begin VB.Image command1 
      Height          =   495
      Left            =   1680
      MouseIcon       =   "frmCommet.frx":030A
      MousePointer    =   99  'Custom
      Top             =   3000
      Width           =   2295
   End
End
Attribute VB_Name = "frmCommet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Nombre As String
Private Sub Command1_Click()

    If Text1 = "" Then
        MsgBox "Debes redactar un mensaje solicitando la paz al lider de " & Nombre
        Exit Sub
    End If

    Call SendData("PEACEOFF" & Nombre & "," & Replace(Text1, vbCrLf, "�"))
    Unload Me

End Sub
Private Sub Command2_Click()

    Unload Me

End Sub
Private Sub Form_Load()

    Me.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\OfertaDePaz.gif")

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bmoving = False And Button = vbLeftButton Then
        Dx3 = X
        dy = Y
        bmoving = True
    End If

End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bmoving And ((X <> Dx3) Or (Y <> dy)) Then Call Move(Left + (X - Dx3), Top + (Y - dy))

End Sub
Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then bmoving = False

End Sub
