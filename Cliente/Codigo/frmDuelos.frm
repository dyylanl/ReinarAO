VERSION 5.00
Begin VB.Form frmDuelos 
   BorderStyle     =   0  'None
   Caption         =   "Duelos 2 vs 2"
   ClientHeight    =   4770
   ClientLeft      =   1005
   ClientTop       =   885
   ClientWidth     =   5985
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmDuelos.frx":0000
   ScaleHeight     =   4770
   ScaleWidth      =   5985
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   5400
      Top             =   3480
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   5400
      Top             =   4200
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3840
      TabIndex        =   3
      Top             =   2280
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   720
      TabIndex        =   2
      Top             =   2280
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   1200
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "20"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000A&
      Height          =   435
      Left            =   3120
      TabIndex        =   4
      Top             =   840
      Width           =   390
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   1320
      Top             =   4080
      Width           =   3495
   End
   Begin VB.Image Image2 
      Height          =   495
      Left            =   1560
      Top             =   3480
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   1080
      Top             =   3000
      Width           =   3975
   End
End
Attribute VB_Name = "frmDuelos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Timer1.Enabled = True
Timer2.Enabled = True
Label1.Caption = "60"
End Sub

Private Sub Image1_Click()
If Text1.Text = "" Or Text2.Text = "" Or Text3.Text = "" Or Text4.Text = "" Then
    MsgBox "Te falta colocar algun nombre de los Participantes"
    Exit Sub
End If
Call SendData("/RETODCT" & " " & Text1.Text)
Call SendData("/RETODCT" & " " & Text2.Text)
Call SendData("/RETODCT" & " " & Text3.Text)
Call SendData("/RETODCT" & " " & Text4.Text)
Timer1.Enabled = False
Timer2.Enabled = False
frmDuelos.Hide
End Sub

Private Sub Image2_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
End Sub

Private Sub Image3_Click()
Call SendData("/RETODCT" & " " & "CANCELARRETOPORRETIRADA")
Timer1.Enabled = False
Timer2.Enabled = False
frmDuelos.Hide
End Sub

Private Sub Timer1_Timer()
Call SendData("/RETODCT" & " " & "CANCELARRETOPORRETIRADA")
AddtoRichTextBox frmMain.rectxt, "Se te acabo el Tiempo de Espera para Realizar este Evento", 2, 51, 223, 1, 1
Timer1.Enabled = False
Timer2.Enabled = False
frmDuelos.Hide
End Sub

Private Sub Timer2_Timer()
Label1.Caption = Label1.Caption - 1
End Sub
