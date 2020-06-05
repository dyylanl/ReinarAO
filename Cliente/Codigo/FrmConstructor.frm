VERSION 5.00
Begin VB.Form FrmConstructor 
   BorderStyle     =   5  'Sizable ToolWindow
   ClientHeight    =   1650
   ClientLeft      =   90
   ClientTop       =   360
   ClientWidth     =   5130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1650
   ScaleWidth      =   5130
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox List2 
      Height          =   645
      ItemData        =   "FrmConstructor.frx":0000
      Left            =   2640
      List            =   "FrmConstructor.frx":000D
      TabIndex        =   4
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Alas Blancas"
      Height          =   255
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   645
      ItemData        =   "FrmConstructor.frx":0055
      Left            =   120
      List            =   "FrmConstructor.frx":0062
      TabIndex        =   2
      Top             =   840
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Alas Rojas"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Requisitos:"
      ForeColor       =   &H00FF0000&
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   5
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Requisitos:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   855
   End
End
Attribute VB_Name = "FrmConstructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
    Call SendData("/ALASBLANCAS")
End Sub

Private Sub Command1_Click()
    Call SendData("/ALASROJAS")
End Sub

