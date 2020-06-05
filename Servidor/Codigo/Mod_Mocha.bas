Attribute VB_Name = "Mod_Mocha"
Option Explicit
Public Type TXY
    x As Integer
    y As Integer
End Type
Public Type TXYTriger
    XY As TXY
    Numero As Integer
End Type
Public Type TXYG1
    XY As TXY
    Numero As Integer
End Type
Public Type TXYG2
    XY As TXY
    Numero As Integer
End Type
Public Type TXYG3
    XY As TXY
    Numero As Integer
End Type
Public Type TXYG4
    XY As TXY
    Numero As Integer
End Type
Public Type TXYObj
    XY As TXY
    mobj As Obj
End Type
Public Type TXYNpc
    XY As TXY
    Numero As Integer
End Type
Public Type TXYSalir
    XY As TXY
    Salida As WorldPos
End Type

Public Sub GrabarSuperMapa(ind As Integer)
    Dim CObj As Integer
    Dim CTriger As Integer
    Dim CG1 As Integer
    Dim CG2 As Integer
    Dim CG3 As Integer
    Dim CG4 As Integer
    Dim CNpc As Integer
    Dim CSalir As Integer
    Dim CBlk As Integer
    
    Dim Mocha_Obj() As TXYObj
    ReDim Mocha_Obj(10001) As TXYObj
    
    Dim Mocha_Triger() As TXYTriger
    ReDim Mocha_Triger(10001) As TXYTriger
    
    Dim Mocha_CG1() As TXYG1
    ReDim Mocha_CG1(10001) As TXYG1
    
    Dim Mocha_CG2() As TXYG2
    ReDim Mocha_CG2(10001) As TXYG2
    
    'Dim Mocha_CG3() As TXYG3
    'ReDim Mocha_CG3(10001) As TXYG3
    
    'Dim Mocha_CG4() As TXYG4
    'ReDim Mocha_CG4(10001) As TXYG4
    
    Dim Mocha_Npc() As TXYNpc
    ReDim Mocha_Npc(10001) As TXYNpc
    
    Dim Mocha_Salir() As TXYSalir
    ReDim Mocha_Salir(10001) As TXYSalir
    
    Dim Mocha_BLK() As TXY
    ReDim Mocha_BLK(10001) As TXY

    Dim x As Integer
    Dim y As Integer
    Dim t As Integer
    Dim tobj As Obj
    Dim tsalir As WorldPos
    For y = YMinMapSize To YMaxMapSize
        For x = XMinMapSize To XMaxMapSize
            tobj = MapData(ind, x, y).OBJInfo
            If tobj.ObjIndex > 0 Then
                Mocha_Obj(CObj).XY.x = x
                Mocha_Obj(CObj).XY.y = y
                Mocha_Obj(CObj).mobj = tobj
                CObj = CObj + 1
            End If
     
            t = MapData(ind, x, y).trigger
            If t > 0 Then
                Mocha_Triger(CTriger).XY.x = x
                Mocha_Triger(CTriger).XY.y = y
                Mocha_Triger(CTriger).Numero = t
                CTriger = CTriger + 1
            End If
            If HayAgua(ind, x, y) Then
                t = MapData(ind, x, y).Graphic(1)
                If t > 0 Then
                    Mocha_CG1(CG1).XY.x = x
                    Mocha_CG1(CG1).XY.y = y
                    Mocha_CG1(CG1).Numero = t
                    CG1 = CG1 + 1
                End If
                
                t = MapData(ind, x, y).Graphic(2)
                If t > 0 Then
                    Mocha_CG2(CG2).XY.x = x
                    Mocha_CG2(CG2).XY.y = y
                    Mocha_CG2(CG2).Numero = t
                    CG2 = CG2 + 1
                End If
            End If
            't = MapData(ind, x, y).Graphic(3)
            'If t > 0 Then
                'Mocha_CG3(CG3).XY.x = x
                'Mocha_CG3(CG3).XY.y = y
                'Mocha_CG3(CG3).Numero = t
                'CG3 = CG3 + 1
            'End If
            
            't = MapData(ind, x, y).Graphic(4)
            'If t > 0 Then
                'Mocha_CG4(CG4).XY.x = x
                'Mocha_CG4(CG4).XY.y = y
                'Mocha_CG4(CG4).Numero = t
                'CG4 = CG4 + 1
            'End If
                        
            
            t = MapData(ind, x, y).NpcIndex
            If t > 0 Then
                Mocha_Npc(CNpc).XY.x = x
                Mocha_Npc(CNpc).XY.y = y
                Mocha_Npc(CNpc).Numero = t
                CNpc = CNpc + 1
            End If
                        
            tsalir = MapData(ind, x, y).TileExit
            If tsalir.Map > 0 Then
                Mocha_Salir(CSalir).XY.x = x
                Mocha_Salir(CSalir).XY.y = y
                Mocha_Salir(CSalir).Salida = tsalir
                CSalir = CSalir + 1
            End If
            t = MapData(ind, x, y).Blocked
            If t > 0 Then
                Mocha_BLK(CBlk).x = x
                Mocha_BLK(CBlk).y = y
                CBlk = CBlk + 1
            End If
        Next
    Next
    ReDim Preserve Mocha_Obj(CObj) As TXYObj
    ReDim Preserve Mocha_Triger(CTriger) As TXYTriger
    ReDim Preserve Mocha_CG1(CG1) As TXYG1
    ReDim Preserve Mocha_CG2(CG2) As TXYG2
    'ReDim Preserve Mocha_CG3(CG3) As TXYG3
    'ReDim Preserve Mocha_CG4(CG4) As TXYG4
    ReDim Preserve Mocha_Npc(CNpc) As TXYNpc
    ReDim Preserve Mocha_Salir(CSalir) As TXYSalir
    ReDim Preserve Mocha_BLK(CBlk) As TXY
    
    Dim ff As Integer
    ff = FreeFile
    Open App.Path & "\Mapitas\" & ind & ".mocha" For Binary Access Write As ff
        Put ff, , CObj
        Put ff, , CTriger
        Put ff, , CG1
        Put ff, , CG2
        'Put ff, , CG3
        'Put ff, , CG4
        Put ff, , CNpc
        Put ff, , CSalir
        Put ff, , CBlk
        
        Put ff, , Mocha_Obj
        Put ff, , Mocha_Triger
        Put ff, , Mocha_CG1
        Put ff, , Mocha_CG2
        'Put ff, , Mocha_CG3
        'Put ff, , Mocha_CG4
        Put ff, , Mocha_Npc
        Put ff, , Mocha_Salir
        Put ff, , Mocha_BLK
    Close ff
End Sub


Public Sub CargarSuperMapa(ind As Integer)
    Dim CObj As Integer
    Dim CTriger As Integer
    Dim CG1 As Integer
    Dim CG2 As Integer
    Dim CG3 As Integer
    Dim CG4 As Integer
    Dim CNpc As Integer
    Dim CSalir As Integer
    Dim CBlk As Integer
    
    Dim Mocha_Obj() As TXYObj
    
    Dim Mocha_Triger() As TXYTriger
    
    Dim Mocha_CG1() As TXYG1
    
    Dim Mocha_CG2() As TXYG2
    
    Dim Mocha_CG3() As TXYG3
    
    Dim Mocha_CG4() As TXYG4
    
    Dim Mocha_Npc() As TXYNpc
    
    Dim Mocha_Salir() As TXYSalir
    
    Dim Mocha_BLK() As TXY
   
    Dim ff As Integer
    ff = FreeFile
    Open App.Path & "\Mapitas\" & ind & ".mocha" For Binary Access Read As ff
        Get ff, , CObj
        Get ff, , CTriger
        Get ff, , CG1
        Get ff, , CG2
        'Get ff, , CG3
        'Get ff, , CG4
        Get ff, , CNpc
        Get ff, , CSalir
        Get ff, , CBlk
    ReDim Preserve Mocha_Obj(CObj) As TXYObj
    ReDim Preserve Mocha_Triger(CTriger) As TXYTriger
    ReDim Preserve Mocha_CG1(CG1) As TXYG1
    ReDim Preserve Mocha_CG2(CG2) As TXYG2
    'ReDim Preserve Mocha_CG3(CG3) As TXYG3
    'ReDim Preserve Mocha_CG4(CG4) As TXYG4
    ReDim Preserve Mocha_Npc(CNpc) As TXYNpc
    ReDim Preserve Mocha_Salir(CSalir) As TXYSalir
    ReDim Preserve Mocha_BLK(CBlk) As TXY
        
        Get ff, , Mocha_Obj
        Get ff, , Mocha_Triger
        Get ff, , Mocha_CG1
        Get ff, , Mocha_CG2
        'Get ff, , Mocha_CG3
        'Get ff, , Mocha_CG4
        Get ff, , Mocha_Npc
        Get ff, , Mocha_Salir
        Get ff, , Mocha_BLK
    Close ff
    Dim t As Integer
    Dim tm As TXYObj
    Dim tt As TXYTriger
    
    Dim tg1 As TXYG1
    Dim tg2 As TXYG2
    Dim tg3 As TXYG3
    Dim tg4 As TXYG4
    
    Dim tnpc As TXYNpc
    Dim tsalir As TXYSalir
    Dim tblk As TXY

    For t = 0 To CObj - 1
        tm = Mocha_Obj(t)
        MapData(ind, tm.XY.x, tm.XY.y).OBJInfo = tm.mobj
    Next
    For t = 0 To CTriger - 1
        tt = Mocha_Triger(t)
        MapData(ind, tt.XY.x, tt.XY.y).trigger = tt.Numero
    Next
    For t = 0 To CG1 - 1
        tg1 = Mocha_CG1(t)
        MapData(ind, tg1.XY.x, tg1.XY.y).Graphic(1) = tg1.Numero
    Next
    For t = 0 To CG2 - 1
        tg2 = Mocha_CG2(t)
        MapData(ind, tg2.XY.x, tg2.XY.y).Graphic(2) = tg2.Numero
    Next
    'For t = 0 To CG3 - 1
        'tg3 = Mocha_CG3(t)
        'MapData(ind, tg3.XY.x, tg3.XY.y).Graphic(3) = tg3.Numero
    'Next
    'For t = 0 To CG4 - 1
        'tg4 = Mocha_CG4(t)
        'MapData(ind, tg4.XY.x, tg4.XY.y).Graphic(4) = tg4.Numero
    'Next
    Dim x As Integer
    Dim y As Integer
    Dim npcfile  As String
    For t = 0 To CNpc - 1
        tnpc = Mocha_Npc(t)
        x = tnpc.XY.x
        y = tnpc.XY.y
        MapData(ind, x, y).NpcIndex = tnpc.Numero
                    If MapData(ind, x, y).NpcIndex > 499 Then
                        npcfile = DatPath & "NPCs-HOSTILES.dat"
                    Else
                        npcfile = DatPath & "NPCs.dat"
                    End If

                    'Si el npc debe hacer respawn en la pos
                    'original la guardamos
                    If val(GetVar(npcfile, "NPC" & MapData(ind, x, y).NpcIndex, "PosOrig")) = 1 Then
                        MapData(ind, x, y).NpcIndex = OpenNPC(MapData(ind, x, y).NpcIndex)
                        Npclist(MapData(ind, x, y).NpcIndex).Orig.Map = ind
                        Npclist(MapData(ind, x, y).NpcIndex).Orig.x = x
                        Npclist(MapData(ind, x, y).NpcIndex).Orig.y = y
                    Else
                        MapData(ind, x, y).NpcIndex = OpenNPC(MapData(ind, x, y).NpcIndex)
                    End If
                            
                    Npclist(MapData(ind, x, y).NpcIndex).Pos.Map = ind
                    Npclist(MapData(ind, x, y).NpcIndex).Pos.x = x
                    Npclist(MapData(ind, x, y).NpcIndex).Pos.y = y
                            
                    Call MakeNPCChar(SendTarget.ToMap, 0, 0, MapData(ind, x, y).NpcIndex, 1, 1, 1)
    Next
    For t = 0 To CSalir - 1
        tsalir = Mocha_Salir(t)
        MapData(ind, tsalir.XY.x, tsalir.XY.y).TileExit = tsalir.Salida
    Next
    For t = 0 To CBlk - 1
        tblk = Mocha_BLK(t)
        MapData(ind, tblk.x, tblk.y).Blocked = 1
    Next
End Sub

