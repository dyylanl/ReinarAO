Attribute VB_Name = "mdlDibujarPersonaje"
    Sub DibujaPJ(Grh As Grh, ByVal x As Integer, ByVal y As Integer, Index As Integer)
    On Error Resume Next
    Dim iGrhIndex As Integer
    If Grh.GrhIndex <= 0 Then Exit Sub
    iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
     
    DrawGrhtoHdc frmCuent.PJ(Index).hDC, iGrhIndex, x, y ' iGrhIndex
    frmCuent.PJ(Index).Refresh
     
    End Sub
    Sub dibujaban(Index As Integer)
    Dim XTexto As Integer
     
    DrawGrhtoHdc frmCuent.PJ(Index).hDC, XTexto, 63, 17030
     
    XTexto = XTexto + 7
    DrawGrhtoHdc frmCuent.PJ(Index).hDC, XTexto, 63, 17029
     
    XTexto = XTexto + 8
    DrawGrhtoHdc frmCuent.PJ(Index).hDC, XTexto, 63, 17042
     
    XTexto = XTexto + 7
    DrawGrhtoHdc frmCuent.PJ(Index).hDC, XTexto, 63, 17042
     
    XTexto = XTexto + 7
    DrawGrhtoHdc frmCuent.PJ(Index).hDC, XTexto, 63, 17033
     
    XTexto = XTexto + 7
    DrawGrhtoHdc frmCuent.PJ(Index).hDC, XTexto, 63, 17032
     
    frmCuent.PJ(Index).Refresh
     
     
    End Sub
     
    Sub dibujamuerto(Index As Integer)
    Dim XTexto As Integer
     
    DrawGrhtoHdc frmCuent.PJ(Index).hDC, XTexto, 0, 17041
     
    XTexto = XTexto + 10
    DrawGrhtoHdc frmCuent.PJ(Index).hDC, XTexto, 0, 17049
     
    XTexto = XTexto + 8
    DrawGrhtoHdc frmCuent.PJ(Index).hDC, XTexto, 0, 17033
     
    XTexto = XTexto + 6
    DrawGrhtoHdc frmCuent.PJ(Index).hDC, XTexto, 0, 17046
     
    XTexto = XTexto + 7
    DrawGrhtoHdc frmCuent.PJ(Index).hDC, XTexto, 0, 17048
     
    XTexto = XTexto + 6
    DrawGrhtoHdc frmCuent.PJ(Index).hDC, XTexto, 0, 17043
     
    frmCuent.PJ(Index).Refresh
    End Sub
    Sub DibujarTodo(ByVal Index As Integer, Body As Integer, Head As Integer, casco As Integer, Shield As Integer, Weapon As Integer, Baned As Integer, Nombre As String, LVL As Integer, Clase As String, muerto As Integer)
     
    Dim Grh As Grh
    Dim Pos As Integer
     
    Dim YBody As Integer
    Dim YYY As Integer
    Dim XBody As Integer
    Dim BBody As Integer
     
    If Baned = 1 Then
        Call dibujaban(Index)
    End If
     
    frmCuent.Nombre(Index).Caption = Nombre
     
    frmCuent.Label1(Index).Caption = "Nivel: " & LVL
    If Clase = 4 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Minero"
    ElseIf Clase = 44 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Bardo"
    ElseIf Clase = 8 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Herrero"
    ElseIf Clase = 14 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Talador"
    ElseIf Clase = 18 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Carpintero"
    ElseIf Clase = 23 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Pescador"
    ElseIf Clase = 27 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Sastre"
    ElseIf Clase = 31 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Alquimista"
    ElseIf Clase = 38 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Mago"
    ElseIf Clase = 39 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Nigromante"
    ElseIf Clase = 41 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Paladin"
    ElseIf Clase = 42 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Clerigo"
    ElseIf Clase = 45 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Druida"
    ElseIf Clase = 47 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Asesino"
    ElseIf Clase = 48 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Cazador"
    ElseIf Clase = 50 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Arquero"
    ElseIf Clase = 51 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Guerrero"
    ElseIf Clase = 56 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Pirata"
    ElseIf Clase = 55 Then
    frmCuent.Label2(Index).Caption = "Clase: " & "Ladron"
    End If
     
    XBody = 8
    YBody = 15
    BBody = 15
     
    If muerto = 1 Then
        Body = 8
        Head = 500
        Shield = 2
        Weapon = 2
        XBody = 10
        YBody = 30
        BBody = 16
        Call dibujamuerto(Index)
    End If
     
    Grh = BodyData(Body).Walk(3)
       
    Call DibujaPJ(Grh, XBody + 4, YBody + 3, Index)
     
    If muerto = 0 Then YYY = BodyData(Body).HeadOffset.y
    If muerto = 1 Then YYY = -9
     
    Pos = YYY + GrhData(GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)).pixelHeight
    Grh = HeadData(Head).Head(3)
     
    If muerto = 1 Then
    Call DibujaPJ(Grh, BBody, Pos - 5, Index)
    Else
    Call DibujaPJ(Grh, BBody + 1, Pos, Index)
    End If
       
    If casco <> 2 And casco > 0 Then
        Grh = CascoAnimData(casco).Head(3)
        Call DibujaPJ(Grh, BBody, Pos, Index)
    End If
     
    If Weapon <> 2 And Weapon > 0 Then
        Grh = WeaponAnimData(Weapon).WeaponWalk(3)
        Call DibujaPJ(Grh, XBody + 1, YBody, Index)
    End If
     
    If Shield <> 2 And Shield > 0 Then
        Grh = ShieldAnimData(Shield).ShieldWalk(3)
        Call DibujaPJ(Grh, XBody, BBody, Index)
    End If
       
    End Sub
