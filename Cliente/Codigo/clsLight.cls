VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Base 0

Private NumLights As Byte

Dim light_last As Long
Public Function Create_Light_To_Map(ByVal map_x As Byte, ByVal map_y As Byte, Optional range As Byte = 3, Optional ByVal Red As Byte = 255, Optional ByVal Green = 255, Optional ByVal Blue As Byte = 255)
    NumLights = NumLights + 1

    ReDim Preserve Lights(1 To NumLights) As Light

    'Le damos color
    Lights(NumLights).RGBcolor.R = Red
    Lights(NumLights).RGBcolor.G = Green
    Lights(NumLights).RGBcolor.b = Blue

    'Alpha (Si borras esto RE KB!!)
    Lights(NumLights).RGBcolor.a = 255

    'Un rango determinado
    Lights(NumLights).range = range

    'La activamos para que se pueda renderizar
    Lights(NumLights).active = True

    'Pos en el mapa
    Lights(NumLights).map_x = map_x
    Lights(NumLights).map_y = map_y

    Call LightRender(NumLights)
End Function
Public Function Delete_Light_To_Map(ByVal X As Byte, ByVal Y As Byte)

    Dim i As Long

    For i = 1 To NumLights
        If Lights(i).map_x = X And Lights(i).map_y = Y Then
            Delete_Light_To_Index i
            Exit Function
        End If
    Next i

End Function
Public Function Delete_Light_To_Index(ByVal light_index As Byte)

    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim Ya As Integer
    Dim Xa As Integer

    Lights(light_index).active = False

    min_x = Lights(light_index).map_x - Lights(light_index).range
    max_x = Lights(light_index).map_x + Lights(light_index).range
    min_y = Lights(light_index).map_y - Lights(light_index).range
    max_y = Lights(light_index).map_y + Lights(light_index).range

    For Ya = min_y To max_y
        For Xa = min_x To max_x
            If InMapBounds(Xa, Ya) Then
                MapData(Xa, Ya).light_value(0) = 0
                MapData(Xa, Ya).light_value(1) = 0
                MapData(Xa, Ya).light_value(2) = 0
                MapData(Xa, Ya).light_value(3) = 0
            End If
        Next Xa
    Next Ya

End Function
Private Function LightCalculate(ByVal cRadio As Integer, ByVal LightX As Integer, ByVal LightY As Integer, ByVal XCoord As Integer, ByVal YCoord As Integer, TileLight As Long, LightColor As D3DCOLORVALUE, AmbientColor As D3DCOLORVALUE) As Long
    Dim XDist As Single
    Dim YDist As Single
    Dim VertexDist As Single
    Dim pRadio As Integer

    Dim CurrentColor As D3DCOLORVALUE

    pRadio = cRadio * 32

    XDist = LightX + 16 - XCoord
    YDist = LightY + 16 - YCoord

    VertexDist = Sqr(XDist * XDist + YDist * YDist)

    If VertexDist <= pRadio Then
        Call D3DXColorLerp(CurrentColor, LightColor, AmbientColor, VertexDist / pRadio)
        LightCalculate = D3DColorXRGB(Round(CurrentColor.R), Round(CurrentColor.G), Round(CurrentColor.b))
        'If TileLight > LightCalculate Then LightCalculate = TileLight
    Else
        LightCalculate = TileLight
    End If
End Function

Private Sub LightRender(ByVal light_index As Integer)

    If light_index = 0 Then Exit Sub
    If Lights(light_index).active = False Then Exit Sub

    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim Color As Long
    Dim Ya As Integer
    Dim Xa As Integer

    Dim TileLight As D3DCOLORVALUE
    Dim AmbientColor As D3DCOLORVALUE
    Dim LightColor As D3DCOLORVALUE

    Dim XCoord As Integer
    Dim YCoord As Integer

    AmbientColor.R = Colorinicial
    AmbientColor.G = Colorinicial
    AmbientColor.b = Colorinicial

    LightColor = Lights(light_index).RGBcolor

    min_x = Lights(light_index).map_x - Lights(light_index).range
    max_x = Lights(light_index).map_x + Lights(light_index).range
    min_y = Lights(light_index).map_y - Lights(light_index).range
    max_y = Lights(light_index).map_y + Lights(light_index).range

    For Ya = min_y To max_y
        For Xa = min_x To max_x
            If InMapBounds(Xa, Ya) Then
                XCoord = Xa * 32
                YCoord = Ya * 32
                MapData(Xa, Ya).light_value(1) = LightCalculate(Lights(light_index).range, Lights(light_index).map_x * 32, Lights(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(1), LightColor, AmbientColor)

                XCoord = Xa * 32 + 32
                YCoord = Ya * 32
                MapData(Xa, Ya).light_value(3) = LightCalculate(Lights(light_index).range, Lights(light_index).map_x * 32, Lights(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(3), LightColor, AmbientColor)

                XCoord = Xa * 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).light_value(0) = LightCalculate(Lights(light_index).range, Lights(light_index).map_x * 32, Lights(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(0), LightColor, AmbientColor)

                XCoord = Xa * 32 + 32
                YCoord = Ya * 32 + 32
                MapData(Xa, Ya).light_value(2) = LightCalculate(Lights(light_index).range, Lights(light_index).map_x * 32, Lights(light_index).map_y * 32, XCoord, YCoord, MapData(Xa, Ya).light_value(2), LightColor, AmbientColor)

            End If
        Next Xa
    Next Ya
End Sub
Public Function Light_Remove(ByVal light_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
'Make sure it's a legal index
    If Light_Check(light_index) Then
        Light_Destroy light_index
        Light_Remove = True
    End If
End Function

Public Function Light_Create(ByVal map_x As Integer, ByVal map_y As Integer, _
                             Optional ByVal range As Byte = 1, Optional ByVal id As Long, Optional ByVal Red As Byte = 255, Optional ByVal Green = 255, Optional ByVal Blue As Byte = 255) As Long
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns the light_index if successful, else 0
'Edited by Juan Martín Sotuyo Dodero
'**************************************************************
    If InMapBounds(map_x, map_y) Then
        'Make sure there is no light in the given map pos
        'If Map_Light_Get(map_x, map_y) <> 0 Then
        '    Light_Create = 0
        '    Exit Function
        'End If
        Light_Create = Light_Next_Open
        Light_Make Light_Create, map_x, map_y, range, id, Red, Green, Blue
    End If
End Function

Public Function Light_Move(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Returns true if successful, else false
'**************************************************************
'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then
        'Make sure it's a legal move
        If InMapBounds(map_x, map_y) Then

            'Move it
            Light_Erase light_index
            Lights(light_index).map_x = map_x
            Lights(light_index).map_y = map_y

            Light_Move = True

        End If
    End If
End Function

Public Function Light_Move_By_Head(ByVal light_index As Long, ByVal Heading As Byte) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 15/05/2002
'Returns true if successful, else false
'**************************************************************
    Dim map_x As Integer
    Dim map_y As Integer
    Dim nX As Integer
    Dim nY As Integer
    Dim addy As Byte
    Dim addx As Byte
    'Check for valid heading
    If Heading < 1 Or Heading > 8 Then
        Light_Move_By_Head = False
        Exit Function
    End If

    'Make sure it's a legal CharIndex
    If Light_Check(light_index) Then

        map_x = Lights(light_index).map_x
        map_y = Lights(light_index).map_y



        Select Case Heading
            Case NORTH
                addy = -1

            Case EAST
                addx = 1

            Case SOUTH
                addy = 1

            Case WEST
                addx = -1
        End Select

        nX = map_x + addx
        nY = map_y + addy

        'Make sure it's a legal move
        If InMapBounds(nX, nY) Then

            'Move it
            Light_Erase light_index

            Lights(light_index).map_x = nX
            Lights(light_index).map_y = nY

            Light_Move_By_Head = True

        End If
    End If
End Function

Private Sub Light_Make(ByVal light_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                       ByVal range As Long, Optional ByVal id As Long, Optional ByVal Red As Byte = 255, Optional ByVal Green = 255, Optional ByVal Blue As Byte = 255)
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
'Update array size
    If light_index > light_last Then
        light_last = light_index
        ReDim Preserve Lights(1 To light_last)
    End If
    light_count = light_count + 1

    'Make active
    Lights(light_index).active = True

    'Le damos color
    Lights(light_index).RGBcolor.R = Red
    Lights(light_index).RGBcolor.G = Green
    Lights(light_index).RGBcolor.b = Blue

    'Alpha (Si borras esto RE KB!!)
    Lights(light_index).RGBcolor.a = 255

    Lights(light_index).map_x = map_x
    Lights(light_index).map_y = map_y
    Lights(light_index).range = range
    Lights(light_index).id = id

End Sub

Private Function Light_Check(ByVal light_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
'check light_index
    If light_index > 0 And light_index <= light_last Then
        If Lights(light_index).active Then
            Light_Check = True
        End If
    End If
End Function
Public Sub Light_Reset_Color(ByVal ColorMuerto As Byte, ByVal Colorinicial As Byte, ByVal EstaMuerto As Boolean)
    Dim loop_counter As Long, a, R, G, b As Byte

    For loop_counter = 1 To light_count
        If Lights(loop_counter).active Then
            '//guardo las originales
            a = Lights(loop_counter).RGBcolor.a
            R = Lights(loop_counter).RGBcolor.R
            G = Lights(loop_counter).RGBcolor.G
            b = Lights(loop_counter).RGBcolor.b
            If EstaMuerto = True Then    '//muere el wachinn
                '//Seteo las luces de la muerte
                Lights(loop_counter).RGBcolor.a = 255
                Lights(loop_counter).RGBcolor.R = ColorMuerto
                Lights(loop_counter).RGBcolor.G = Colorinicial
                Lights(loop_counter).RGBcolor.b = Colorinicial
            Else
                Lights(loop_counter).RGBcolor.a = 255
                Lights(loop_counter).RGBcolor.R = R
                Lights(loop_counter).RGBcolor.G = G
                Lights(loop_counter).RGBcolor.b = b
            End If
        End If
    Next loop_counter

End Sub
Public Sub Light_Render_All()
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim loop_counter As Long

    For loop_counter = 1 To light_count

        If Lights(loop_counter).active Then
            LightRender loop_counter
        End If

    Next loop_counter
End Sub

Private Function Light_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    On Error GoTo ErrorHandler:
    Dim LoopC As Long

    LoopC = 1
    Do Until Lights(LoopC).active = False
        If LoopC = light_last Then
            Light_Next_Open = light_last + 1
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop

    Light_Next_Open = LoopC
    Exit Function
ErrorHandler:
    Light_Next_Open = 1
End Function

Public Function Light_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
    On Error GoTo ErrorHandler:
    Dim LoopC As Long

    LoopC = 1
    Do Until Lights(LoopC).id = id
        If LoopC = light_last Then
            Light_Find = 0
            Exit Function
        End If
        LoopC = LoopC + 1
    Loop

    Light_Find = LoopC
    Exit Function
ErrorHandler:
    Light_Find = 0
End Function

Public Function Light_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long

    For Index = 1 To light_last
        'Make sure it's a legal index
        If Light_Check(Index) Then
            Light_Destroy Index
        End If
    Next Index

    Light_Remove_All = True
End Function

Private Sub Light_Destroy(ByVal light_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim Temp As Light

    Light_Erase light_index

    Lights(light_index) = Temp

    'Update array size
    If light_index = light_last Then
        Do Until Lights(light_last).active
            light_last = light_last - 1
            If light_last = 0 Then
                light_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve Lights(1 To light_last)
    End If
    light_count = light_count - 1
End Sub

Private Sub Light_Erase(ByVal light_index As Long)
'***************************************'
'Author: Juan Martín Sotuyo Dodero
'Last modified: 3/31/2003
'Correctly erases a light
'***************************************'
    Dim min_x As Integer
    Dim min_y As Integer
    Dim max_x As Integer
    Dim max_y As Integer
    Dim X As Integer
    Dim Y As Integer

    'Set up light borders
    min_x = Lights(light_index).map_x - Lights(light_index).range
    min_y = Lights(light_index).map_y - Lights(light_index).range
    max_x = Lights(light_index).map_x + Lights(light_index).range
    max_y = Lights(light_index).map_y + Lights(light_index).range

    'Arrange corners
    'NE
    If InMapBounds(min_x, min_y) Then
        MapData(min_x, min_y).light_value(2) = 0
    End If
    'NW
    If InMapBounds(max_x, min_y) Then
        MapData(max_x, min_y).light_value(0) = 0
    End If
    'SW
    If InMapBounds(max_x, max_y) Then
        MapData(max_x, max_y).light_value(1) = 0
    End If
    'SE
    If InMapBounds(min_x, max_y) Then
        MapData(min_x, max_y).light_value(3) = 0
    End If

    'Arrange borders
    'Upper border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, min_y) Then
            MapData(X, min_y).light_value(0) = 0
            MapData(X, min_y).light_value(2) = 0
        End If
    Next X

    'Lower border
    For X = min_x + 1 To max_x - 1
        If InMapBounds(X, max_y) Then
            MapData(X, max_y).light_value(1) = 0
            MapData(X, max_y).light_value(3) = 0
        End If
    Next X

    'Left border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(min_x, Y) Then
            MapData(min_x, Y).light_value(2) = 0
            MapData(min_x, Y).light_value(3) = 0
        End If
    Next Y

    'Right border
    For Y = min_y + 1 To max_y - 1
        If InMapBounds(max_x, Y) Then
            MapData(max_x, Y).light_value(0) = 0
            MapData(max_x, Y).light_value(1) = 0
        End If
    Next Y

    'Set the inner part of the light
    For X = min_x + 1 To max_x - 1
        For Y = min_y + 1 To max_y - 1
            If InMapBounds(X, Y) Then
                MapData(X, Y).light_value(0) = 0
                MapData(X, Y).light_value(1) = 0
                MapData(X, Y).light_value(2) = 0
                MapData(X, Y).light_value(3) = 0
            End If
        Next Y
    Next X
End Sub











