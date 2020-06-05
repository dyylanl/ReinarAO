VERSION 5.00
Begin VB.Form frmListaAmigos 
   BorderStyle     =   0  'None
   Caption         =   "Lista de Amigos de Genius AO"
   ClientHeight    =   4245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmListaAmigos.frx":0000
   ScaleHeight     =   4245
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   1630
      Left            =   3060
      TabIndex        =   2
      Top             =   1750
      Width           =   1150
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   285
      Left            =   360
      TabIndex        =   1
      Top             =   390
      Width           =   2415
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000004&
      Height          =   2370
      ItemData        =   "frmListaAmigos.frx":13165
      Left            =   360
      List            =   "frmListaAmigos.frx":13167
      TabIndex        =   0
      Top             =   960
      Width           =   2415
   End
   Begin VB.Image Image4 
      Height          =   135
      Left            =   0
      MouseIcon       =   "frmListaAmigos.frx":13169
      MousePointer    =   99  'Custom
      Top             =   4080
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   615
      Left            =   3000
      MouseIcon       =   "frmListaAmigos.frx":13473
      MousePointer    =   99  'Custom
      Top             =   240
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   615
      Left            =   3000
      MouseIcon       =   "frmListaAmigos.frx":1377D
      MousePointer    =   99  'Custom
      Top             =   960
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   3000
      MouseIcon       =   "frmListaAmigos.frx":13A87
      MousePointer    =   99  'Custom
      Top             =   3600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000012&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000014&
      Height          =   195
      Left            =   1560
      TabIndex        =   3
      Top             =   3720
      Width           =   765
   End
End
Attribute VB_Name = "frmListaAmigos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Picture = LoadPicture(DirGraficos & "ListaAmigos.jpg")
Call SendData("LISFF" & frmMain.Label8)
End Sub
Private Sub Image1_Click()
Call SendData("/MP" & Chr(64) & List1.Text & Chr(64) & Text2.Text)
End Sub
Private Sub Image2_Click()
If List1.ListIndex = -1 Then
MsgBox "Debes seleccionar algun campo de la lista"
Exit Sub
ElseIf List1.ListIndex = 10 Then
MsgBox "Has llegado al limite de amigos"
Exit Sub
Else
Call SendData("DELFF" & Chr(64) & List1.Text & Chr(64) & List1.ListIndex)

End If
End Sub

Private Sub Image3_Click()
If List1.ListIndex = -1 Then
MsgBox "Debes seleccionar algun campo de la lista"
Exit Sub
ElseIf List1.ListIndex = 10 Then
MsgBox "Has llegado al limite de amigos"
Exit Sub
Else
Call SendData("ADDFF" & Chr(64) & Text1.Text & Chr(64) & List1.ListIndex)
End If
End Sub

Private Sub Image4_Click()
Unload Me
frmMain.SetFocus
End Sub

Private Sub list1_Click()
SendData ("ESTFF" & List1.Text)
End Sub
