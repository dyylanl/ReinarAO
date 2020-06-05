VERSION 5.00
Begin VB.Form Torneo2 
   BackColor       =   &H00000000&
   Caption         =   "Torneo 2vs2."
   ClientHeight    =   5115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5745
   LinkTopic       =   "Form1"
   ScaleHeight     =   5115
   ScaleWidth      =   5745
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command13 
      Caption         =   "2 - 1 A Favor"
      Height          =   375
      Left            =   3120
      TabIndex        =   19
      Top             =   3000
      Width           =   2535
   End
   Begin VB.CommandButton Command12 
      Caption         =   "2 - 1 A Favor"
      Height          =   375
      Left            =   120
      TabIndex        =   18
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Pierden"
      Height          =   375
      Left            =   3120
      TabIndex        =   16
      Top             =   3480
      Width           =   2535
   End
   Begin VB.CommandButton Command10 
      Caption         =   "2 - 0 A Favor"
      Height          =   375
      Left            =   3120
      TabIndex        =   15
      Top             =   2520
      Width           =   2535
   End
   Begin VB.CommandButton Command9 
      Caption         =   "1 - 1"
      Height          =   375
      Left            =   3120
      TabIndex        =   14
      Top             =   2040
      Width           =   2535
   End
   Begin VB.CommandButton Command8 
      Caption         =   "1 - 0 A Favor"
      Height          =   375
      Left            =   3120
      TabIndex        =   13
      Top             =   1560
      Width           =   2535
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Ganan Torneo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      TabIndex        =   12
      Top             =   4560
      Width           =   2775
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   3240
      TabIndex        =   11
      Text            =   "Nick Personaje"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1440
      TabIndex        =   10
      Text            =   "Nick Personaje"
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Pierden"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   3480
      Width           =   2655
   End
   Begin VB.CommandButton Command5 
      Caption         =   "2 - 0 A Favor"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   2655
   End
   Begin VB.CommandButton Command4 
      Caption         =   "1 - 1"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   2040
      Width           =   2655
   End
   Begin VB.CommandButton Command3 
      Caption         =   "1 - 0 A Favor"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Duelean 2º Vez"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   5535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Duelean"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   5535
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   1440
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Y"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2880
      TabIndex        =   20
      Top             =   4080
      Width           =   255
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "VS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2760
      TabIndex        =   17
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "Torneo2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
    Call SendData("/RMSG Torneo 2VS2> Se enfrentan en esta nueva Ronda:" & " " & Text1.Text & " Y " & Text2.Text & " vs " & Text3.Text & " Y " & Text4.Text)
    Call SendData("/RMSG Torneo 2VS2> Suerte para Ambas Parejas..")
    Call SendData("/RMSG Torneo 2VS2> Esquinas y sale en..")
    Call SendData("/CUENTA 10")
    Call SendData("/TELEP" & " " & Text1.Text & " " & "14 39 21")
    Call SendData("/TELEP" & " " & Text2.Text & " " & "14 40 21")
    Call SendData("/TELEP" & " " & Text3.Text & " " & "14 53 34")
    Call SendData("/TELEP" & " " & Text4.Text & " " & "14 52 34")
End Sub

Private Sub Command10_Click()
    Call SendData("/RMSG Torneo 2VS2> 2 A 0 A favor de" & " " & Text3.Text & " Y " & Text4.Text)
    Call SendData("/REVIVIR" & " " & Text1.Text)
    Call SendData("/REVIVIR" & " " & Text2.Text)
End Sub

Private Sub Command11_Click()
    Call SendData("/RMSG Torneo 2VS2> Pierden:" & " " & Text3.Text & " Y " & Text4.Text & " . Quedan descalificados del Torneo.")
    Call SendData("/TELEP" & " " & Text3.Text & " " & "1 50 50")
    Call SendData("/TELEP" & " " & Text4.Text & " " & "1 50 50")
End Sub

Private Sub Command12_Click()
    Call SendData("/RMSG Torneo 2VS2> 2 A 1 A favor de" & " " & Text1.Text & " Y " & Text2.Text)
    Call SendData("/REVIVIR" & " " & Text3.Text)
    Call SendData("/REVIVIR" & " " & Text4.Text)
End Sub

Private Sub Command13_Click()
    Call SendData("/RMSG Torneo 2VS2> 2 A 1 A favor de" & " " & Text3.Text & " Y " & Text4.Text)
    Call SendData("/REVIVIR" & " " & Text1.Text)
    Call SendData("/REVIVIR" & " " & Text2.Text)
End Sub

Private Sub Command2_Click()
    Call SendData("/RMSG Torneo 2VS2> Se enfrentan Nuevamente:" & " " & Text1.Text & " Y " & Text2.Text & " vs " & Text3.Text & " Y " & Text4.Text)
    Call SendData("/RMSG Torneo 2VS2> Suerte para Ambas Parejas..")
    Call SendData("/RMSG Torneo 2VS2> Esquinas y sale en..")
    Call SendData("/CUENTA 10")
    Call SendData("/TELEP" & " " & Text1.Text & " " & "14 39 21")
    Call SendData("/TELEP" & " " & Text2.Text & " " & "14 40 21")
    Call SendData("/TELEP" & " " & Text3.Text & " " & "14 53 34")
    Call SendData("/TELEP" & " " & Text4.Text & " " & "14 52 34")
End Sub

Private Sub Command3_Click()
    Call SendData("/RMSG Torneo 2VS2> 1 A 0 A favor de" & " " & Text1.Text & " Y " & Text2.Text)
    Call SendData("/REVIVIR" & " " & Text3.Text)
    Call SendData("/REVIVIR" & " " & Text4.Text)
End Sub
Private Sub Command4_Click()
    Call SendData("/RMSG Torneo 2VS2> Lo empatan" & " " & Text1.Text & " Y " & Text2.Text)
    Call SendData("/REVIVIR" & " " & Text3.Text)
    Call SendData("/REVIVIR" & " " & Text4.Text)
End Sub

Private Sub Command5_Click()
    Call SendData("/RMSG Torneo 2VS2> 2 A 0 A favor de" & " " & Text1.Text & " Y " & Text2.Text)
    Call SendData("/REVIVIR" & " " & Text3.Text)
    Call SendData("/REVIVIR" & " " & Text4.Text)
End Sub

Private Sub Command6_Click()
    Call SendData("/RMSG Torneo 2VS2> Pierden:" & " " & Text1.Text & " Y " & Text2.Text & " . Quedan descalificados del Torneo.")
    Call SendData("/TELEP" & " " & Text1.Text & " " & "1 50 50")
    Call SendData("/TELEP" & " " & Text2.Text & " " & "1 50 50")
End Sub

Private Sub Command7_Click()
    Call SendData("/RMSG Torneo 2VS2> Y los Ganadores del Torneo son:" & " " & Text5.Text & " Y " & Text6.Text)
    Call SendData("/RMSG Gracias por Participar..! y Felicitaciones a los Ganadores!!")
End Sub

Private Sub Command8_Click()
    Call SendData("/RMSG Torneo> 1 A 0 A favor de" & " " & Text3.Text & " Y " & Text4.Text)
    Call SendData("/REVIVIR" & " " & Text1.Text)
    Call SendData("/REVIVIR" & " " & Text2.Text)
End Sub

Private Sub Command9_Click()
    Call SendData("/RMSG Torneo 2VS2> Lo empatan" & " " & Text3.Text & " Y " & Text4.Text)
    Call SendData("/REVIVIR" & " " & Text1.Text)
    Call SendData("/REVIVIR" & " " & Text2.Text)
End Sub

