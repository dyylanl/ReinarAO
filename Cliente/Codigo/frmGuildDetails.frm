VERSION 5.00
Begin VB.Form frmGuildDetails 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Detalles del Clan"
   ClientHeight    =   7650
   ClientLeft      =   0
   ClientTop       =   -105
   ClientWidth     =   7635
   ClipControls    =   0   'False
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
   ScaleHeight     =   7650
   ScaleWidth      =   7635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00111720&
      Height          =   285
      Index           =   0
      Left            =   960
      TabIndex        =   8
      Top             =   3240
      Width           =   5775
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   1
      Left            =   960
      TabIndex        =   7
      Top             =   3600
      Width           =   5775
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   2
      Left            =   960
      TabIndex        =   6
      Top             =   3960
      Width           =   5775
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   3
      Left            =   960
      TabIndex        =   5
      Top             =   4320
      Width           =   5775
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   4
      Left            =   960
      TabIndex        =   4
      Top             =   4680
      Width           =   5775
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   5
      Left            =   960
      TabIndex        =   3
      Top             =   5040
      Width           =   5775
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   6
      Left            =   960
      TabIndex        =   2
      Top             =   5400
      Width           =   5775
   End
   Begin VB.TextBox txtCodex1 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   285
      Index           =   7
      Left            =   960
      TabIndex        =   1
      Top             =   5760
      Width           =   5775
   End
   Begin VB.TextBox txtDesc 
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00111720&
      Height          =   1575
      Left            =   700
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   840
      Width           =   6375
   End
   Begin VB.Image Command1 
      Height          =   375
      Index           =   0
      Left            =   120
      Top             =   7200
      Width           =   975
   End
   Begin VB.Image Command1 
      Height          =   615
      Index           =   1
      Left            =   5400
      Top             =   6720
      Width           =   1935
   End
End
Attribute VB_Name = "frmGuildDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Command1_Click(Index As Integer)
    Select Case Index

        Case 0
            Unload Me
        Case 1
            Dim fdesc$
            fdesc$ = Replace(txtDesc, vbCrLf, "º", , , vbBinaryCompare)






            Dim k As Integer
            Dim Cont As Integer
            Cont = 0
            For k = 0 To txtCodex1.UBound




                If Len(txtCodex1(k).Text) > 0 Then Cont = Cont + 1
            Next

            If Cont < 4 Then
                MsgBox "Debes definir al menos cuatro mandamientos."
                Exit Sub
            End If

            Dim chunk As String

            If CreandoClan Then
                chunk = "CIG" & fdesc$ & "¬" & ClanName & "¬" & Site
            Else
                chunk = "DESCOD" & fdesc$
            End If

            chunk = chunk & "¬"

            For k = 0 To Cont - 1
                chunk = chunk & txtCodex1(k) & "|"
            Next

            Call SendData(Left$(chunk, Len(chunk) - 1))

            CreandoClan = False

            Unload Me

    End Select



End Sub

Private Sub Form_Deactivate()

    If Not frmGuildLeader.Visible Then
        Me.SetFocus
    Else

    End If


End Sub

Private Sub Form_Load()
    Me.Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\GuildDetailsCodex.gif")

End Sub
