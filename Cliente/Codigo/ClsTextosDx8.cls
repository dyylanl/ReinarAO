VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsTextosDx8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory _
                     Lib "kernel32" _
                         Alias "RtlMoveMemory" (ByRef Destination As Any, _
                                                ByRef Source As Any, _
                                                ByVal Length As Long)

'Describes the return from a texture init
Private Type D3DXIMAGE_INFO_A

    Width As Long
    Height As Long

End Type

Private Type POINTAPI

    X As Long
    Y As Long

End Type

Private Type CharVA

    Vertex(0 To 3) As TLVERTEX

End Type

Private Type VFH

    BitmapWidth As Long         'Size of the bitmap itself
    BitmapHeight As Long
    CellWidth As Long           'Size of the cells (area for each character)
    CellHeight As Long
    BaseCharOffset As Byte      'The character we start from
    CharWidth(0 To 255) As Byte    'The actual factual width of each character
    CharVA(0 To 255) As CharVA

End Type

Private Type CustomFont

    HeaderInfo As VFH           'Holds the header information
    Texture As Direct3DTexture8    'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As POINTAPI     'Size of the texture

End Type

'Private Const Font_Default_TextureNum As Long = -1   'The texture number used to represent this font - only used for AlternateRendering - keep negative to prevent interfering with game textures
Private cfonts(1 To 2) As CustomFont    ' _Default2 As CustomFont

Public Sub Text_Draw(ByVal Left As Long, _
                     ByVal Top As Long, _
                     ByVal Text As String, _
                     ByVal Color As Long, _
                     Optional ByVal center As Boolean = False, _
                     Optional ByVal Alpha As Byte = 255, _
                     Optional size As Integer = 1)

    If Alpha <> 255 Then

        Dim newRGB As D3DCOLORVALUE

        ARGBtoD3DCOLORVALUE Color, newRGB
        Color = D3DColorARGB(Alpha, newRGB.R, newRGB.G, newRGB.b)
    End If

    Text_Render cfonts(1), Text, Left, Top, Color, center, size

End Sub


Private Sub Text_Render(ByRef UseFont As CustomFont, _
                        ByVal Text As String, _
                        ByVal X As Long, _
                        ByVal Y As Long, _
                        ByVal Color As Long, _
                        ByVal center As Boolean, _
                        Optional size As Integer = 1)

    Dim TempVA(0 To 3) As TLVERTEX

    Dim tempstr() As String

    Dim Count As Integer

    Dim ascii() As Byte

    Dim i As Long

    Dim j As Long

    'Dim KeyPhrase As Byte
    Dim tempcolor As Long

    Dim ResetColor As Byte

    Dim YOffset As Single

    'Check if we have the device
    If d3ddevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub

    'Get the text into arrays (split by vbCrLf)
    tempstr = Split(Text, vbCrLf)

    'Set the temp color (or else the first character has no color)
    tempcolor = Color

    'Set the texture
    d3ddevice.SetTexture 0, UseFont.Texture

    If center Then
        X = X - Text_GetWidth(Text) * 0.5
    End If

    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(tempstr)

        If Len(tempstr(i)) > 0 Then
            YOffset = i * UseFont.CharHeight
            Count = 0

            'Convert the characters to the ascii value
            ascii() = StrConv(tempstr(i), vbFromUnicode)

            'Loop through the characters
            For j = 1 To Len(tempstr(i))

                'Copy from the cached vertex array to the temp vertex array
                CopyMemory TempVA(0), UseFont.HeaderInfo.CharVA(ascii(j - 1)).Vertex(0), 32 * 4

                'Set up the verticies
                TempVA(0).X = X + Count * size
                TempVA(0).Y = Y + YOffset

                TempVA(1).X = TempVA(1).X + X + Count * size
                TempVA(1).Y = TempVA(0).Y

                TempVA(2).X = TempVA(0).X
                TempVA(2).Y = TempVA(2).Y + TempVA(0).Y * size

                TempVA(3).X = TempVA(1).X
                TempVA(3).Y = TempVA(2).Y

                'Set the colors
                TempVA(0).Color = tempcolor
                TempVA(1).Color = tempcolor
                TempVA(2).Color = tempcolor
                TempVA(3).Color = tempcolor

                d3ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, TempVA(0), Len(TempVA(0))

                'Shift over the the position to render the next character
                Count = Count + UseFont.HeaderInfo.CharWidth(ascii(j - 1))

                'End If

                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    tempcolor = Color
                End If

            Next j

        End If

    Next i

End Sub

Public Function Text_GetWidth(ByVal Text As String) As Integer
'***************************************************
'Returns the width of text
'More info: http://www.vbgore.com/GameClient.TileEn ... tTextWidth
'***************************************************
    Dim i As Long

    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function

    'Loop through the text
    For i = 1 To Len(Text)

        'Add up the stored character widths
        Text_GetWidth = Text_GetWidth + cfonts(1).HeaderInfo.CharWidth(Asc(mid$(Text, i, 1)))

    Next i

End Function

Sub Text_Init_Textures()

'*****************************************************************
'Init the custom font textures
'More info: http://www.vbgore.com/GameClient.TileEn ... ntTextures
'*****************************************************************
    Dim TexInfo As D3DXIMAGE_INFO_A

    On Error GoTo eDebug:

    'Check if we have the device
    If d3ddevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    'Set the texture
    Set cfonts(1).Texture = D3DX.CreateTextureFromFileEx(d3ddevice, App.Path & "\RECURSOS\Graficos\grh\Fuente_default.png", D3DX_DEFAULT, D3DX_DEFAULT, D3DX_DEFAULT, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_POINT, &HFF000000, ByVal 0, ByVal 0)

    'Store the size of the texture
    cfonts(1).TextureSize.X = TexInfo.Width
    cfonts(1).TextureSize.Y = TexInfo.Height

    Exit Sub

eDebug:

    If Err.Number = "-2005529767" Then
        MsgBox "Error en la textura utilizada de DirectX 8", vbCritical

        End

    End If

    End

End Sub

Sub Text_Init_Settings()

    Dim FileNum As Byte

    Dim LoopChar As Long

    Dim Row As Single

    Dim u As Single

    Dim V As Single

    '*** Default font ***

    'Load the header information
    FileNum = FreeFile
    Open App.Path & "\RECURSOS\INIT\fuente_default.dat" For Binary As #FileNum
    Get #FileNum, , cfonts(1).HeaderInfo
    Close #FileNum

    'Calculate some common values
    cfonts(1).CharHeight = cfonts(1).HeaderInfo.CellHeight - 4
    cfonts(1).RowPitch = cfonts(1).HeaderInfo.BitmapWidth \ cfonts(1).HeaderInfo.CellWidth
    cfonts(1).ColFactor = cfonts(1).HeaderInfo.CellWidth / cfonts(1).HeaderInfo.BitmapWidth
    cfonts(1).RowFactor = cfonts(1).HeaderInfo.CellHeight / cfonts(1).HeaderInfo.BitmapHeight

    'Cache the verticies used to draw the character (only requires setting the color and adding to the X/Y values)
    For LoopChar = 0 To 255

        'tU and tV value (basically tU = BitmapXPosition / BitmapWidth, and height for tV)
        Row = (LoopChar - cfonts(1).HeaderInfo.BaseCharOffset) \ cfonts(1).RowPitch
        u = ((LoopChar - cfonts(1).HeaderInfo.BaseCharOffset) - (Row * cfonts(1).RowPitch)) * cfonts(1).ColFactor
        V = Row * cfonts(1).RowFactor

        'Set the verticies
        With cfonts(1).HeaderInfo.CharVA(LoopChar)
            .Vertex(0).Color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            .Vertex(0).rhw = 1
            .Vertex(0).tu = u
            .Vertex(0).tv = V
            .Vertex(0).X = 0
            .Vertex(0).Y = 0
            .Vertex(0).Z = 0

            .Vertex(1).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(1).rhw = 1
            .Vertex(1).tu = u + cfonts(1).ColFactor
            .Vertex(1).tv = V
            .Vertex(1).X = cfonts(1).HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).Z = 0

            .Vertex(2).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(2).rhw = 1
            .Vertex(2).tu = u
            .Vertex(2).tv = V + cfonts(1).RowFactor
            .Vertex(2).X = 0
            .Vertex(2).Y = cfonts(1).HeaderInfo.CellHeight
            .Vertex(2).Z = 0

            .Vertex(3).Color = D3DColorARGB(255, 0, 0, 0)
            .Vertex(3).rhw = 1
            .Vertex(3).tu = u + cfonts(1).ColFactor
            .Vertex(3).tv = V + cfonts(1).RowFactor
            .Vertex(3).X = cfonts(1).HeaderInfo.CellWidth
            .Vertex(3).Y = cfonts(1).HeaderInfo.CellHeight
            .Vertex(3).Z = 0
        End With

    Next LoopChar

End Sub

Public Sub Text_Destroy()

'end texts, its necesary? GDK
    Erase cfonts()

End Sub






