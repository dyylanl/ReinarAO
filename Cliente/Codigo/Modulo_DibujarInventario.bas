Attribute VB_Name = "modInventarioGrafico"
'FénixAO en DX8 by ·Parra, Thusing y DarkTester

Option Explicit

Public Const XCantItems = 5

Public OffsetDelInv As Integer
Public ItemElegido As Integer
Public mx As Integer
Public my As Integer


Sub ActualizarOtherInventory(Slot As Integer)

    If OtherInventory(Slot).OBJIndex = 0 Then
        frmComerciar.List1(0).List(Slot - 1) = "Nada"
    Else
        frmComerciar.List1(0).List(Slot - 1) = OtherInventory(Slot).Name
    End If

    If frmComerciar.List1(0).ListIndex = Slot - 1 And lista = 0 Then Call ActualizarInformacionComercio(0)

End Sub
Sub ActualizarInventario(Slot As Integer)

    If UserInventory(Slot).Amount = 0 Then
        frmMain.imgObjeto(Slot).ToolTipText = "Nada"
        frmMain.lblObjCant(Slot).ToolTipText = "Nada"
        frmMain.lblObjCant(Slot).Caption = ""
        If ItemElegido = Slot Then frmMain.Shape1.Visible = False
    Else
        frmMain.imgObjeto(Slot).ToolTipText = UserInventory(Slot).Name
        frmMain.lblObjCant(Slot).ToolTipText = UserInventory(Slot).Name
        frmMain.lblObjCant(Slot).Caption = CStr(UserInventory(Slot).Amount)
        If ItemElegido = Slot Then frmMain.Shape1.Visible = True
    End If

    If UserInventory(Slot).GrhIndex > 0 Then
        If Extract_File(Graphics, App.Path & "\RECURSOS\GRAFICOS\", GrhData(UserInventory(Slot).GrhIndex).FileNum & ".png", App.Path & "\RECURSOS\GRAFICOS\") Then
            Call PngImageLoad(DirGraficos & GrhData(UserInventory(Slot).GrhIndex).FileNum & ".png", frmMain.imgObjeto(Slot))
            'frmMain.imgObjeto(Slot).picture = LoadPicture(DirGraficos & GrhData(UserInventory(Slot).GrhIndex).FileNum & ".png")
            Call Kill(DirGraficos & GrhData(UserInventory(Slot).GrhIndex).FileNum & ".png")
        End If
    Else
        frmMain.imgObjeto(Slot).Picture = LoadPicture()
    End If


    If UserInventory(Slot).Equipped > 0 Then
        frmMain.Label2(Slot).Visible = True
    Else
        frmMain.Label2(Slot).Visible = False
    End If

    If frmComerciar.Visible Then
        If UserInventory(Slot).Amount = 0 Then
            frmComerciar.List1(1).List(Slot - 1) = "Nada"
        Else
            frmComerciar.List1(1).List(Slot - 1) = UserInventory(Slot).Name
        End If
        If frmComerciar.List1(1).ListIndex = Slot - 1 And lista = 1 Then Call ActualizarInformacionComercio(1)
    End If

End Sub
