VERSION 5.00
Begin VB.Form frmProcesos 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   ClientHeight    =   8805
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5505
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8805
   ScaleWidth      =   5505
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   480
      TabIndex        =   4
      Text            =   "Usuario"
      Top             =   8280
      Width           =   1575
   End
   Begin VB.Timer Timer2 
      Left            =   2880
      Top             =   7440
   End
   Begin VB.Timer Timer1 
      Left            =   2400
      Top             =   7440
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Kick Proceso"
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   7440
      Width           =   1575
   End
   Begin VB.ListBox List1 
      Height          =   6105
      ItemData        =   "frmProcesos.frx":0000
      Left            =   120
      List            =   "frmProcesos.frx":0002
      TabIndex        =   0
      Top             =   240
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmProcesos.frx":0004
      ForeColor       =   &H000000FF&
      Height          =   795
      Left            =   480
      TabIndex        =   3
      Top             =   6480
      Width           =   4545
   End
End
Attribute VB_Name = "frmProcesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Declare Function CloseHandle Lib "Kernel32" (ByVal hObject _
                                                     As Long) As Long


Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
Const WS_EX_LAYERED = &H80000
Private Declare Function GetWindowLong Lib "USER32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "USER32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "USER32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Private ret As Long
Private cont2 As Integer

Sub RellenaLista()
    Dim hSnapshot As Long
    Dim uProceso As PROCESSENTRY32
    Dim res As Long
    List1.Clear
    hSnapshot = CreateToolhelpSnapshot(2&, 0&)
    If hSnapshot <> 0 Then
        uProceso.dwSize = Len(uProceso)
        res = ProcessFirst(hSnapshot, uProceso)
        Do While res
            List1.AddItem Left$(uProceso.szexeFile, InStr(uProceso.szexeFile, Chr$(0)) - 1)
            List1.ItemData(List1.NewIndex) = uProceso.th32ProcessID
            res = ProcessNext(hSnapshot, uProceso)
        Loop
        Call CloseHandle(hSnapshot)
    End If
End Sub
Private Sub Command1_Click()
    Call SendData("/Killprocess" & " " & Text1.Text & " " & List1.Text)
End Sub

Private Sub Command2_Click()
    Cancel = True
    Timer2.Enabled = True
    Unload Me
End Sub

Private Sub Form_Load()
    cont2 = 255
    ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED
    SetWindowLong Me.hWnd, GWL_EXSTYLE, ret
    Timer1.Interval = 1
    Timer2.Interval = 1
    Timer2.Enabled = False
    Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
    Static Cont As Integer
    Cont = Cont + 5
    If Cont > 255 Then
        Cont = 0
        Timer1.Enabled = False
    Else
        SetLayeredWindowAttributes Me.hWnd, 0, Cont, LWA_ALPHA
    End If
End Sub

Private Sub Timer2_Timer()
    cont2 = cont2 - 5
    If cont2 < 0 Then
        Timer2.Enabled = False
        End
    Else
        SetLayeredWindowAttributes Me.hWnd, 0, cont2, LWA_ALPHA
    End If
End Sub
