VERSION 5.00
Begin VB.Form frmCambiarPasswd 
   BorderStyle     =   0  'None
   Caption         =   "Cambiar password"
   ClientHeight    =   3390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox ConfirPasswdNuevo 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "$"
      TabIndex        =   2
      Top             =   2160
      Width           =   1455
   End
   Begin VB.TextBox PasswdNuevo 
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   3360
      PasswordChar    =   "$"
      TabIndex        =   1
      Top             =   1800
      Width           =   1455
   End
   Begin VB.TextBox PasswdViejo 
      Height          =   285
      Left            =   3360
      TabIndex        =   0
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Image Cancelar 
      Height          =   255
      Left            =   4800
      MouseIcon       =   "frmCambiarPassword.frx":0000
      MousePointer    =   99  'Custom
      Top             =   3120
      Width           =   615
   End
   Begin VB.Image Aceptar 
      Height          =   495
      Left            =   1080
      MouseIcon       =   "frmCambiarPassword.frx":030A
      MousePointer    =   99  'Custom
      Top             =   2760
      Width           =   1575
   End
End
Attribute VB_Name = "frmCambiarPasswd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Aceptar_Click()

    If Me.PasswdNuevo <> Me.ConfirPasswdNuevo Then
        Call MsgBox("El password nuevo no coincide con su confirmaci�n.")
        Exit Sub
    End If

    If Len(Me.PasswdNuevo) < 6 Then
        Call AddtoRichTextBox(frmMain.rectxt, "El password nuevo debe tener al menos 6 caracteres.", 65, 190, 156, False, False, False)
        Exit Sub
    End If

    Call SendData("PASS" & MD5String(Me.PasswdViejo) & "," & MD5String(Me.PasswdNuevo))

    Unload Me

End Sub

Private Sub Cancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    Me.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\password.gif")

End Sub
