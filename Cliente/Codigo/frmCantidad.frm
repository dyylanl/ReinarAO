VERSION 5.00
Begin VB.Form frmCantidad 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   1890
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3540
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
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command4 
      Caption         =   "TIRAR TODO"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Tirar"
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H004DC488&
      Height          =   375
      Left            =   480
      MaxLength       =   7
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Escribe la cantidad"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   960
      TabIndex        =   1
      Top             =   240
      Width           =   1605
   End
End
Attribute VB_Name = "frmCantidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command3_Click()
    frmCantidad.Visible = False
    Call SendData("TI" & ItemElegido & "," & frmCantidad.Text1.Text)
    frmCantidad.Text1.Text = "0"
End Sub

Private Sub Command4_Click()
    frmCantidad.Visible = False

    If ItemElegido <> FLAGORO Then
        Call SendData("TI" & ItemElegido & "," & UserInventory(ItemElegido).Amount)
    Else: Call SendData("TI" & ItemElegido & "," & UserGLD)
    End If

    frmCantidad.Text1.Text = "0"
End Sub

Private Sub Form_Deactivate()

    Unload Me

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
Private Sub Text1_Change()

    If Val(Text1.Text) < 0 Then
        Text1.Text = MAX_INVENTORY_OBJS
    End If

    If Val(Text1.Text) > MAX_INVENTORY_OBJS And ItemElegido <> FLAGORO Then
        Text1.Text = 1
    End If

End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)

    If (KeyAscii <> 8) Then
        If (Index <> 6) And (KeyAscii < 48 Or KeyAscii > 57) Then KeyAscii = 0
    End If

End Sub

