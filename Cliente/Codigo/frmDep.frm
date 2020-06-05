VERSION 5.00
Begin VB.Form frmDep 
   BorderStyle     =   0  'None
   ClientHeight    =   2220
   ClientLeft      =   1755
   ClientTop       =   2850
   ClientWidth     =   5385
   LinkTopic       =   "Form1"
   ScaleHeight     =   2220
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtDatos 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty DataFormat 
         Type            =   0
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   0
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   220
      Left            =   240
      TabIndex        =   0
      Top             =   1050
      Width           =   4875
   End
   Begin VB.Label lblDatos 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1920
      TabIndex        =   1
      Top             =   600
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   2880
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   240
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "frmDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    Me.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\DepositarN.bmp")
End Sub

Private Sub Image2_Click()
    Unload Me
End Sub

Private Sub Image1_Click()
    If Val(txtDatos.Text) <= 0 Then
        lblDatos.Caption = "Cantidad inválida."
        Exit Sub
    End If


    Call SendData("/DEPOSITAR " & Val(txtDatos.Text))

    Unload Me

End Sub
