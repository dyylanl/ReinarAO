VERSION 5.00
Begin VB.Form frmCentroConsulta 
   Caption         =   "Form1"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   9630
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtRespuesta 
      Height          =   1335
      Left            =   3000
      TabIndex        =   7
      Text            =   "Text1"
      Top             =   3600
      Width           =   5655
   End
   Begin VB.TextBox txtConsulta 
      Height          =   2295
      Left            =   3000
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   600
      Width           =   5535
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Enviar Respuesta."
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   5280
      Width           =   4455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Borrar Todo"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Borrar Consulta"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   4800
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "/SUM"
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   4200
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "/IRA"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   3600
      Width           =   1695
   End
   Begin VB.ListBox lstUsuarios 
      Height          =   2595
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   2055
   End
   Begin VB.Image sendResp 
      Height          =   375
      Left            =   3600
      Top             =   5280
      Width           =   4455
   End
   Begin VB.Image Image4 
      Height          =   375
      Index           =   3
      Left            =   360
      Top             =   5400
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   375
      Index           =   2
      Left            =   360
      Top             =   4800
      Width           =   1695
   End
   Begin VB.Image Image2 
      Height          =   375
      Index           =   1
      Left            =   360
      Top             =   4200
      Width           =   1695
   End
   Begin VB.Image Image1 
      Height          =   375
      Index           =   0
      Left            =   360
      Top             =   3600
      Width           =   1695
   End
End
Attribute VB_Name = "frmCentroConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub imgOption_Click(Index As Integer)
Select Case Index
    Case 0
        SendData ("/IRA " & ReadField(1, lstUsuarios.List(lstUsuarios.ListIndex), Asc("-")))
    Case 1
        SendData ("/SUM " & ReadField(1, lstUsuarios.List(lstUsuarios.ListIndex), Asc("-")))
    Case 2
        If lstUsuarios.ListIndex < 0 Then Exit Sub
        SendData ("BORRACONSULTA" & lstUsuarios.List(lstUsuarios.ListIndex))
        lstUsuarios.RemoveItem lstUsuarios.ListIndex
    Case 3
        Call SendData("/BORRARSOS")
        lstUsuarios.Clear
End Select
 
End Sub
 
Private Sub lstUsuarios_Click()
Dim ind As Integer
ind = Val(ReadField(2, lstUsuarios.List(lstUsuarios.ListIndex), Asc("-")))
SendData ("/VERCONSULTA " & ReadField(1, lstUsuarios.List(lstUsuarios.ListIndex), Asc("-")))
End Sub
Private Sub sendResp_Click()
SendData ("/RESPONDER " & txtConsulta.Text & "@" & ReadField(1, lstUsuarios.List(lstUsuarios.ListIndex), Asc("-")))
End Sub
 
 
