VERSION 5.00
Begin VB.Form frmbp 
   BorderStyle     =   0  'None
   ClientHeight    =   4065
   ClientLeft      =   2760
   ClientTop       =   3435
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   Picture         =   "frmbp.frx":0000
   ScaleHeight     =   4065
   ScaleWidth      =   4605
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image5 
      Height          =   495
      Left            =   960
      Top             =   2880
      Width           =   2655
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   960
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   960
      Top             =   1440
      Width           =   2655
   End
   Begin VB.Image Image2 
      Height          =   255
      Left            =   960
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   960
      Top             =   600
      Width           =   2655
   End
End
Attribute VB_Name = "frmbp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Image1_Click()
    For i = 1 To UBound(UserInventory)
        frmComerciar.List1(1).AddItem UserInventory(i).Name
    Next
    frmComerciar.Image2(0).Left = 182
    frmComerciar.cantidad.Left = 248
    frmComerciar.Image2(1).Visible = False
    frmComerciar.precio.Visible = False
    frmComerciar.Image1(0).Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Retirar.gif")
    frmComerciar.Image1(1).Picture = LoadPicture(App.Path & "\RECURSOS\INTERFACES\Depositar.gif")

    Comerciando = 2
    frmComerciar.Show
End Sub

Private Sub Image3_Click()
    frmDep.Show
End Sub

Private Sub Image4_Click()
    frmRet.Show
End Sub

Private Sub Image5_Click()
    Unload Me
End Sub
