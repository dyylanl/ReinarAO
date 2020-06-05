Attribute VB_Name = "Mod_TileEngine"
Option Explicit

Private bTechoAB As Byte
Private AlphaNPC As Byte
Public SurfaceDB As clsSurfaceManDynDX8

Public ScrollPixelsPerFrame As Double
Public USUARI0 As Integer
Public Const PI As Single = 3.14159265358979

Public WeatherFogX1 As Single
Public WeatherFogY1 As Single
Public WeatherFogX2 As Single
Public WeatherFogY2 As Single
Public WeatherDoFog As Byte
Public WeatherFogCount As Byte
Public EndTime As Long
Private Const ScreenWidth As Long = 800


Const HASH_TABLE_SIZE As Long = 337

Private mD3D As D3DX8
Private device As Direct3DDevice8

' Parra was here (;
Private MaxMemory As Long
Private mGraphicsNumber As Long
Private mCurrentMemoryBytes As Long
Private mMaxMemoryBytes As Long
Private Const DEFAULT_MEMORY_TO_USE As Long = 64    ' In MB
Private Const BYTES_PER_MB As Long = 1048576

Private Type SURFACE_ENTRY_DYN
    filename As Integer
    UltimoAcceso As Long
    Texture As Direct3DTexture8
    size As Long
    texture_width As Integer
    texture_height As Integer
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private TexList(HASH_TABLE_SIZE - 1) As HashNode

Private lFrameLimiter As Long
Public lFrameModLimiter As Long
Public lFrameTimer As Long
Public timerTicksPerFrame As Double    'mmmm me encanta que sea Double jaja
Public timerElapsedTime As Single
Public engineBaseSpeed As Single

'Describes a transformable lit vertex
Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    rhw As Single
    Color As Long
    Specular As Long
    tu As Single
    tv As Single
End Type

'********** Direct X ***********
Private Type D3D8Textures
    Texture As Direct3DTexture8
    texwidth As Long
    TexHeight As Long
    Dimension As Integer
End Type


'DirectX 8 Objects
Public dX As DirectX8
Public D3D As Direct3D8
Public D3DX As D3DX8
Public d3ddevice As Direct3DDevice8

'Font List
Dim font_count As Long
Dim font_last As Long

Public font_list() As D3DXFont

Public Enum FontAlignment
    fa_center = DT_CENTER
    fa_top = DT_TOP
    fa_left = DT_LEFT
    fa_topleft = DT_TOP Or DT_LEFT
    fa_bottomleft = DT_BOTTOM Or DT_LEFT
    fa_bottom = DT_BOTTOM
    fa_right = DT_RIGHT
    fa_bottomright = DT_BOTTOM Or DT_RIGHT
    fa_topright = DT_TOP Or DT_RIGHT
End Enum

Public mFreeMemoryBytes As Long

Private pUdtMemStatus As MEMORYSTATUS

Private Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Private Declare Sub GlobalMemoryStatus Lib "Kernel32" (lpBuffer As MEMORYSTATUS)
Private Declare Sub CopyMemory _
                     Lib "Kernel32" _
                         Alias "RtlMoveMemory" (ByRef Destination As Any, _
                                                ByRef Source As Any, _
                                                ByVal Length As Long)
                                                
Public Declare Sub CopyMemory2 Lib "Kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public base_light As Long
Public day_r_old As Byte
Public day_g_old As Byte
Public day_b_old As Byte

Type luzxhora

    R As Long
    G As Long
    b As Long

End Type


Public luz_dia(0 To 24) As luzxhora    '¬¬ la hora 24 dura 1 minuto entre las 24 y las 0

Public Const ImgSize As Byte = 4

Public Const XMaxMapSize = 100
Public Const XMinMapSize = 1
Public Const YMaxMapSize = 100
Public Const YMinMapSize = 1

Public Const GrhFogata = 1521


Public Const SRCCOPY = &HCC0020

Public Type Position
    X As Integer
    Y As Integer
End Type

Public Type Position2
    X As Double
    Y As Double
End Type

Public Type WorldPos
    map As Integer
    X As Integer
    Y As Integer
End Type

Public Type GrhData
    sX As Integer
    sY As Integer
    FileNum As Integer
    pixelWidth As Integer
    pixelHeight As Integer
    TileWidth As Single
    TileHeight As Single

    NumFrames As Integer
    Frames(1 To 25) As Integer
    speed As Integer
    active As Boolean
End Type

Public Type Grh
    GrhIndex As Integer
    FrameCounter As Double
    SpeedCounter As Byte
    Started As Byte
End Type

Public Type BodyData
    Walk(1 To 4) As Grh
    HeadOffset As Position
End Type

Public Type HeadData
    Head(1 To 4) As Grh
End Type

Type WeaponAnimData
    WeaponWalk(1 To 4) As Grh
End Type

Type ShieldAnimData
    ShieldWalk(1 To 4) As Grh
End Type

Public Type FxData
    FX As Grh
    OffSetX As Long
    OffSetY As Long
End Type

Public Type Char
    Alas As BodyData
    ParticleIndex As Integer
    Aura_index As Long
    Aura_Angle As Single
    Particula As Integer
    active As Byte
    Heading As Byte
    Pos As Position
    Body As BodyData
    Head As HeadData
    casco As HeadData
    arma As WeaponAnimData
    escudo As ShieldAnimData
    UsandoArma As Boolean
    FX As Integer
    FxLoopTimes As Integer
    Criminal As Byte
    Navegando As Byte
    EsPremium As Byte
    Nombre As String

    haciendoataque As Byte
    Moving As Byte
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    MoveOffset As Position2
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    moved As Boolean

End Type

Public Type Obj
    OBJIndex As Integer
    Amount As Integer
End Type

Public Type MapBlock
    parti_index As Integer
    particle_group_index As Integer
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    Trigger As Byte

'    base_light(0 To 3) As Boolean    'Indica si el tile tiene luz propia.
    light_index As Integer
    'light_base_value(0 To 3) As Long    'Luz propia del tile.
    light_value(0 To 3) As Long    'Color de luz con el que esta siendo renderizado.
    '//Sangre VBGore
    Blood As Byte
    Ambient As String
End Type

Public IniPath As String
Public MapPath As String

Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

Public CurMap As Integer
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position
Public AddtoUserPos As Position
Public UserCharIndex As Integer

Public EngineRun As Boolean
Public FramesPerSec As Integer
Public FramesPerSecCounter As Long

Public WindowTileWidth As Integer
Public WindowTileHeight As Integer

Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

Public LastChar As Integer

Public GrhData() As GrhData
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As FxData
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
Public Grh() As Grh
Public MapData() As MapBlock
Public CharList(1 To 10000) As Char

Public bRain As Boolean    'está raineando?
Public bTecho As Boolean    'hay techo?
Public brstTick As Long

Private LTLluvia(4) As Integer

Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
    plFogata = 3
End Enum
'RGB Type
Public Type RGB
    R As Long
    G As Long
    b As Long
End Type

Dim map_current As map

Private Type decoration
    Grh As Grh
    Render_On_Top As Boolean
    subtile_pos As Byte
End Type

Private Type Map_Tile
    Grh(1 To 3) As Grh
    decoration(1 To 5) As decoration
    decoration_count As Byte
    Blocked As Boolean
    particle_group_index As Long
    char_index As Long
    'light_base_value(0 To 3) As Long
    light_value(0 To 3) As Long

    exit_index As Long
    npc_index As Long
    item_index As Long

    Trigger As Byte
End Type

Private Type map
    map_grid() As Map_Tile
    map_x_max As Long
    map_x_min As Long
    map_y_max As Long
    map_y_min As Long

End Type



Rem Mannakia .. Parituclas ORE 1.0.

Private Type Particle
    TimeAlpha As Single
    Alpha As Single
    friction As Single
    X As Single
    Y As Single
    vector_x As Single
    vector_y As Single
    angle As Single
    Grh As Grh
    alive_counter As Long
    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Integer
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    Rgb_List As D3DCOLORVALUE
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type

Dim base_tile_size As Integer

'Modified by: Ryan Cain (Onezero)
'Last modify date: 5/14/2003
Private Type particle_group
    active As Boolean
    id As Long
    map_x As Integer
    map_y As Integer
    char_index As Long

    frame_counter As Single
    frame_speed As Single

    stream_type As Byte

    particle_stream() As Particle
    particle_count As Long

    grh_index_list() As Long
    grh_index_count As Long

    alpha_blend As Boolean

    alive_counter As Long
    never_die As Boolean

    X1 As Integer
    X2 As Integer
    Y1 As Integer
    Y2 As Integer
    angle As Integer
    vecx1 As Integer
    vecx2 As Integer
    vecy1 As Integer
    vecy2 As Integer
    life1 As Long
    life2 As Long
    fric As Long
    spin_speedL As Single
    spin_speedH As Single
    gravity As Boolean
    grav_strength As Long
    bounce_strength As Long
    spin As Boolean
    XMove As Boolean
    YMove As Boolean
    move_x1 As Integer
    move_x2 As Integer
    move_y1 As Integer
    move_y2 As Integer
    Rgb_List(3) As Long

    'Added by Juan Martín Sotuyo Dodero
    speed As Single
    life_counter As Long

    'Added by David Justus
    grh_resize As Boolean
    grh_resizex As Integer
    grh_resizey As Integer
End Type
'Particle system

'Dim StreamData() As particle_group

Dim particle_group_list() As particle_group
Public particle_group_count As Long
Dim particle_group_last As Long
Rem mannakia
'BitBlt

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "Kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "Kernel32" (lpPerformanceCount As Currency) As Long
Private Function AmigoClan(ByVal CharIndex As Integer) As Boolean
    Dim Nombre1 As String
    Dim Nombre2 As String

    Nombre1 = CharList(UserCharIndex).Nombre
    Nombre2 = CharList(CharIndex).Nombre

    If InStr(Nombre1, "<") > 0 And InStr(Nombre2, "<") > 0 Then

        AmigoClan = Trim$(mid$(Nombre2, InStr(Nombre2, "<"))) = _
                    Trim$(mid$(Nombre1, InStr(Nombre1, "<")))
    End If
End Function
Public Function GetElapsedTime() As Single
    Dim start_time As Currency
    Static end_time As Currency
    Static timer_freq As Currency

    'Get the timer frequency
    If timer_freq = 0 Then
        QueryPerformanceFrequency timer_freq
    End If

    'Get current time
    Call QueryPerformanceCounter(start_time)

    'Calculate elapsed time
    GetElapsedTime = (start_time - end_time) / timer_freq * 1000

    'Get next end time
    Call QueryPerformanceCounter(end_time)
End Function

Public Sub ShowNextFrame()


    Dim ulttick As Long, esttick As Long
    Dim timers(1 To 5) As Long
    Dim loopc As Long

    Do While prgRun
        'If Not EngineRun Then
            'Effect_Die_Connect
        '    RenderConnect
        'End If
        Call RefreshAllChars
        If EngineRun Then
            If frmMain.WindowState <> 1 Then

                If UserMoving Then
                    '****** Move screen Left and Right if needed ******
                    If AddtoUserPos.X <> 0 Then
                        OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrame * AddtoUserPos.X * timerTicksPerFrame
                        If Abs(OffsetCounterX) >= Abs(32 * AddtoUserPos.X) Then
                            OffsetCounterX = 0
                            AddtoUserPos.X = 0
                            UserMoving = False
                        End If
                    End If

                    '****** Move screen Up and Down if needed ******
                    If AddtoUserPos.Y <> 0 Then
                        OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrame * AddtoUserPos.Y * timerTicksPerFrame
                        If Abs(OffsetCounterY) >= Abs(32 * AddtoUserPos.Y) Then
                            OffsetCounterY = 0
                            AddtoUserPos.Y = 0
                            UserMoving = False
                        End If
                    End If
                End If

                d3ddevice.BeginScene
                d3ddevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, D3DColorXRGB(0, 0, 0), 1#, 0


                If UserCiego Then
                    d3ddevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
                Else
                    RenderScreen UserPos.X - AddtoUserPos.X, UserPos.Y - AddtoUserPos.Y, OffsetCounterX, OffsetCounterY
                End If

                Call Texto.Text_Draw(10, 10, "FPS: " & FramesPerSec & ".", D3DColorXRGB(225, 255, 255))
                If Niebla = True Then Texto.Text_Draw 10, 25, "Neblina", D3DColorXRGB(255, 255, 255)

                If Cartel Then DibujarCartel
                If Dialogos.CantidadDialogos <> 0 Then Dialogos.MostrarTexto
                'RenderSounds
                Effect_Die_Connect
                d3ddevice.Present RenderRect, ByVal 0, frmMain.renderer.hWnd, ByVal 0

                d3ddevice.EndScene

                lFrameLimiter = GetTickCount
                FramesPerSecCounter = FramesPerSecCounter + 1
                timerElapsedTime = GetElapsedTime()
                timerTicksPerFrame = timerElapsedTime * engineBaseSpeed

                'FramesPerSecCounter = FramesPerSecCounter + 1
            End If
            If (Opciones.Audio = 1 Or Opciones.sMusica <> CONST_DESHABILITADA) Then Call Sound.Sound_Render
        End If
        
        If GetTickCount - lFrameTimer > 1000 Then
            FramesPerSec = FramesPerSecCounter
            'If FramesPerSec <> 0 Then ScrollPixelsPerFrame = 150 / FramesPerSec
            FramesPerSecCounter = 0
            lFrameTimer = GetTickCount
        End If
        
        
        If Not Pausa And frmMain.Visible And Not frmForo.Visible Then
            CheckKeys
        ElseIf frmConnect.Visible = True Then
            If Not UserMap = 1 Then
                UserMap = 1
                
                SwitchMapNew UserMap, True
                SwitchMapNew UserMap, False
            End If
            RenderConnect
        End If

        

        If Opciones.FPSConfig = 1 Then    '18 FPS
            While (GetTickCount - lFrameTimer) \ 64 < FramesPerSecCounter
                Sleep 5
            Wend
        ElseIf Opciones.FPSConfig = 2 Then    '32 FPS
            ActualizarBarras
            While (GetTickCount - lFrameTimer) \ 34 < FramesPerSecCounter
                Sleep 5
            Wend
        ElseIf Opciones.FPSConfig = 3 Then    '64 FPS
            ActualizarBarras
            While (GetTickCount - lFrameTimer) \ 16 < FramesPerSecCounter
                Sleep 5
            Wend
        ElseIf Opciones.FPSConfig = 4 Then    'libres
            ActualizarBarras
        End If


        ' ### I N T E R V A L O S ###
        esttick = GetTickCount
        'For loopc = 1 To UBound(timers)
        For loopc = 1 To UBound(timers)
            timers(loopc) = timers(loopc) + (esttick - ulttick)

            If timers(1) >= tUs Then
                timers(1) = 0
                NoPuedeUsar = False
            End If
        Next loopc
        ulttick = GetTickCount

        DoEvents
    Loop

End Sub
'AURASSSSSSSS
Private Sub Draw_Aura(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal center As Byte, ByVal Animate As Byte, Optional ByVal Alpha As Boolean, Optional ByVal angle As Single, Optional Aura_index As Long)
    On Error Resume Next
    Dim CurrentGrhIndex As Integer
    Dim Light(3) As Long

    If Grh.GrhIndex = 0 Then Exit Sub

    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.SpeedCounter)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1

                If Grh.FrameCounter <> -1 Then
                    If Grh.FrameCounter > 0 Then
                        Grh.FrameCounter = Grh.FrameCounter - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If

    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

    'Center Grh over X,Y pos
    If center Then

        If GrhData(CurrentGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(CurrentGrhIndex).TileWidth * (32 \ 2)) + 32 \ 2
        End If

        If GrhData(Grh.GrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(CurrentGrhIndex).TileHeight * 32) + 32
        End If

    End If

    Light(0) = D3DColorXRGB(Aura(Aura_index).R, Aura(Aura_index).G, Aura(Aura_index).b)
    Light(1) = Light(0)
    Light(2) = Light(0)
    Light(3) = Light(0)

    Device_Box_Textured_Render_Advance CurrentGrhIndex, _
                                       X, Y, _
                                       GrhData(CurrentGrhIndex).pixelWidth, GrhData(CurrentGrhIndex).pixelHeight, _
                                       Light(), GrhData(CurrentGrhIndex).sX, GrhData(CurrentGrhIndex).sY, _
                                       Alpha _
                                       , angle

End Sub

Sub DDrawTechoGrhtoSurface(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0, Optional ByVal map_x As Byte, Optional ByVal map_y As Byte, Optional Alpha As Boolean, Optional ByVal angle As Single, Optional AlphaB As Boolean = False, Optional alphaa As Byte = 255)
    On Error Resume Next
    Dim iGrhIndex As Integer
    Dim QuitarAnimacion As Boolean


    If Animate Then
        If Grh.Started = 1 Then
            If Grh.SpeedCounter > 0 Then
                Grh.SpeedCounter = Grh.SpeedCounter - 1
                If Grh.SpeedCounter = 0 Then
                    Grh.SpeedCounter = GrhData(Grh.GrhIndex).speed
                    Grh.FrameCounter = Grh.FrameCounter + (1 / (8 / ScrollPixelsPerFrame))
                    If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                        Grh.FrameCounter = 1
                        If KillAnim Then
                            If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then

                                If CharList(KillAnim).FxLoopTimes > 0 Then CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                                If CharList(KillAnim).FxLoopTimes < 1 Then
                                    CharList(KillAnim).FX = 0
                                    Exit Sub
                                End If

                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    If Grh.GrhIndex = 0 Then Exit Sub


    iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)


    If center Then
        If GrhData(iGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16
        End If
        If GrhData(iGrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32
        End If
    End If

    If map_x Or map_y = 0 Then map_x = 1: map_y = 1

    Device_Box_Textured_Render_Advance iGrhIndex, _
                                       X, Y, _
                                       GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
                                       MapData(map_x, map_y).light_value, _
                                       GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
                                       Alpha, angle, AlphaB, , , alphaa


End Sub

Sub DDrawTransGrhtoSurface(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, ByRef Color() As Long, Optional ByVal KillAnim As Integer = 0, Optional ByVal map_x As Byte, Optional ByVal map_y As Byte, Optional Alpha As Boolean, Optional ByVal angle As Single, Optional AlphaB As Boolean = False, Optional alphaa As Byte = 255)
On Error Resume Next
    Dim iGrhIndex As Integer


If Animate Then
    If Grh.Started = 1 Then
       
        Grh.FrameCounter = Grh.FrameCounter + ((timerElapsedTime * 0.1) * GrhData(Grh.GrhIndex).NumFrames / Grh.SpeedCounter)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
               
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                   
                If KillAnim <> 0 Then
                    If CharList(KillAnim).FX > 0 Then
                        If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                          CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes <= 0 Then CharList(KillAnim).FX = 0: Exit Sub
                        End If
                    End If
                End If
            End If
    End If
End If

        If Grh.GrhIndex = 0 Then Exit Sub
    
    
        iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    
        If center Then
            If GrhData(iGrhIndex).TileWidth <> 1 Then
                X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16
            End If
            If GrhData(iGrhIndex).TileHeight <> 1 Then
                Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32
            End If
        End If
    
        If map_x Or map_y = 0 Then map_x = 1: map_y = 1
    
        Device_Box_Textured_Render_Advance iGrhIndex, _
                                           X, Y, _
                                           GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
                                           Color(), _
                                           GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
                                           Alpha, angle, AlphaB, , , alphaa
    

End Sub
Sub DrawGrhtoHdc(picX As PictureBox, Grh As Integer, ByVal X As Integer, ByVal Y As Integer)
    On Error Resume Next
    Dim hdcsrc As Long
    Dim file_path As String

    If Grh <= 0 Then Exit Sub

    'If it's animated switch GrhIndex to first frame
    If GrhData(Grh).NumFrames <> 1 Then
        Grh = GrhData(Grh).Frames(1)
    End If

    If Extract_File(Graphics, App.Path & "\Recursos\GRAFICOS\", GrhData(Grh).FileNum & ".png", App.Path & "\Recursos\GRAFICOS\") Then
        file_path = App.Path & "\Recursos\GRAFICOS\" & GrhData(Grh).FileNum & ".png"
        Call PngPictureLoad(file_path, picX, False)
        'Call PngPictureLoad(file_path, frmBancoObj.Picture1, False)
    End If

    Call Kill(App.Path & "\Recursos\Graficos\*.png")
End Sub
Sub RenderScreen(ByVal TileX As Integer, ByVal TileY As Integer, ByVal PixelOffsetX As Double, ByVal PixelOffsetY As Double)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 28/04/2010
'Last modified by: Franco (Thusing)
'Renders everything to the viewport
'**************************************************************
    'On Error Resume Next
    If UserCiego Then Exit Sub
    Dim Y As Integer                'Keeps track of where on map we are
    Dim X As Integer                'Keeps track of where on map we are
    Dim minY As Integer             'Start Y pos on current map
    Dim maxY As Integer             'End Y pos on current map
    Dim minX As Integer             'Start X pos on current map
    Dim maxX As Integer             'End X pos on current map
    Dim ScreenX As Integer             'Keeps track of where to place tile on screen
    Dim ScreenY As Integer             'Keeps track of where to place tile on screen
    Dim minXOffset As Integer
    Dim minYOffset As Integer
    Dim PixelOffsetXTemp As Integer    'For centering grhs
    Dim PixelOffsetYTemp As Integer    'For centering grhs
    Dim TempChar As Char
    Dim moved As Byte
    Dim iPPx As Integer
    Dim iPPy As Integer

    'Figure out Ends and Starts of screen
    ScreenMinY = TileY - HalfWindowTileHeight
    ScreenMaxY = TileY + HalfWindowTileHeight
    ScreenMinX = TileX - HalfWindowTileWidth
    ScreenMaxX = TileX + HalfWindowTileWidth

    minY = ScreenMinY - TileBufferSize
    maxY = ScreenMaxY + TileBufferSize
    minX = ScreenMinX - TileBufferSize
    maxX = ScreenMaxX + TileBufferSize

    'Make sure mins and maxs are allways in map bounds
    If minY < XMinMapSize Then
        minYOffset = YMinMapSize - minY
        minY = YMinMapSize
    End If

    If maxY > YMaxMapSize Then maxY = YMaxMapSize

    If minX < XMinMapSize Then
        minXOffset = XMinMapSize - minX
        minX = XMinMapSize
    End If

    If maxX > XMaxMapSize Then maxX = XMaxMapSize

    'If we can, we render around the view area to make it smoother
    If ScreenMinY > YMinMapSize Then
        ScreenMinY = ScreenMinY - 1
    Else
        ScreenMinY = 1
        ScreenY = 1
    End If

    If ScreenMaxY < YMaxMapSize Then ScreenMaxY = ScreenMaxY + 1

    If ScreenMinX > XMinMapSize Then
        ScreenMinX = ScreenMinX - 1
    Else
        ScreenMinX = 1
        ScreenX = 1
    End If

    If ScreenMaxX < XMaxMapSize Then ScreenMaxX = ScreenMaxX + 1
    ParticleOffsetX = (Engine_PixelPosX(ScreenMinX) - PixelOffsetX)
    ParticleOffsetY = (Engine_PixelPosY(ScreenMinY) - PixelOffsetY)

    'Draw floor layer
    For Y = ScreenMinY To ScreenMaxY
        For X = ScreenMinX To ScreenMaxX
            'Layer 1 **********************************
            Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(1), _
                                        (ScreenX - 1) * 32 + PixelOffsetX, _
                                        (ScreenY - 1) * 32 + PixelOffsetY, _
                                        0, 1, MapData(X, Y).light_value, , X, Y)
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(2), (ScreenX - 1) * 32 + PixelOffsetX, (ScreenY - 1) * 32 + PixelOffsetY, 1, 1, MapData(X, Y).light_value)
            End If
            '******************************************
            ScreenX = ScreenX + 1
        Next X

        'Reset ScreenX to original value and increment ScreenY
        ScreenX = ScreenX - X + ScreenMinX
        ScreenY = ScreenY + 1
    Next Y

    '//Sangre VBGore
    d3ddevice.SetVertexShader FVF2

    Engine_Render_Blood

    d3ddevice.SetVertexShader FVF


    ScreenY = minYOffset - TileBufferSize
    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            PixelOffsetXTemp = ScreenX * 32 + PixelOffsetX
            PixelOffsetYTemp = ScreenY * 32 + PixelOffsetY
            With MapData(X, Y)
                '*****************************************
                'Object Layer **********************************
                If .ObjGrh.GrhIndex <> 0 Then

                    Call DDrawTransGrhtoSurface(.ObjGrh, _
                        PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, MapData(X, Y).light_value)
                End If





                'Renderizado del personaje ************************
                If MapData(X, Y).CharIndex > 0 Then
                    TempChar = CharList(MapData(X, Y).CharIndex)
                    PixelOffsetXTemp = PixelOffsetX
                    PixelOffsetYTemp = PixelOffsetY
                    moved = False

                    'If needed, move left and right
                    With TempChar
                        If .Moving Then
                        'If needed, move left and right
                            If .scrollDirectionX <> 0 Then
                                .MoveOffset.X = .MoveOffset.X + ScrollPixelsPerFrame * Sgn(.scrollDirectionX) * timerTicksPerFrame
                            
                                'Start animations
                                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                                If .Body.Walk(.Heading).SpeedCounter > 0 Then _
                                    .Body.Walk(.Heading).Started = 1
                                .arma.WeaponWalk(.Heading).Started = 1
                                .escudo.ShieldWalk(.Heading).Started = 1
                                .Alas.Walk(.Heading).Started = 1
                            
                                'Char moved
                                moved = True
                            
                                'Check if we already got there
                                If (Sgn(.scrollDirectionX) = 1 And .MoveOffset.X >= 0) Or _
                                    (Sgn(.scrollDirectionX) = -1 And .MoveOffset.X <= 0) Then
                                    .MoveOffset.X = 0
                                    .scrollDirectionX = 0
                                End If
                            End If
                        
                            'If needed, move up and down
                            If .scrollDirectionY <> 0 Then
                                .MoveOffset.Y = .MoveOffset.Y + ScrollPixelsPerFrame * Sgn(.scrollDirectionY) * timerTicksPerFrame
                            
                                'Start animations
                                'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                                If .Body.Walk(.Heading).SpeedCounter > 0 Then _
                                    .Body.Walk(.Heading).Started = 1
                                .arma.WeaponWalk(.Heading).Started = 1
                                .escudo.ShieldWalk(.Heading).Started = 1
                                .Alas.Walk(.Heading).Started = 1
                            
                                'Char moved
                                moved = True
                            
                            'Check if we already got there
                                If (Sgn(.scrollDirectionY) = 1 And .MoveOffset.Y >= 0) Or _
                                    (Sgn(.scrollDirectionY) = -1 And .MoveOffset.Y <= 0) Then
                                    .MoveOffset.Y = 0
                                    .scrollDirectionY = 0
                                End If
                            End If
                        End If
             
                        If .Heading = 0 Then .Heading = 3
             
                        If moved = 0 Then
                            .Body.Walk(.Heading).Started = 0
                            .Body.Walk(.Heading).FrameCounter = 1
                           
                            .arma.WeaponWalk(.Heading).Started = 0
                            .arma.WeaponWalk(.Heading).FrameCounter = 1
                           
                            .escudo.ShieldWalk(.Heading).Started = 0
                            .escudo.ShieldWalk(.Heading).FrameCounter = 1
                           
                            .Alas.Walk(.Heading).Started = 0
                            .Alas.Walk(.Heading).FrameCounter = 1
                            .Moving = 0
                        End If
                       
                        If TempChar.haciendoataque = 0 And .MoveOffset.X = 0 And .MoveOffset.Y = 0 Then
                            .arma.WeaponWalk(.Heading).Started = 0
                            '.arma.WeaponWalk(.Heading).FrameCounter = 1
                            .escudo.ShieldWalk(.Heading).Started = 0
                           
                            End If
                           
                        If TempChar.haciendoataque = 1 Then
                            .arma.WeaponWalk(.Heading).Started = 1
                            .escudo.ShieldWalk(.Heading).Started = 1
                            .arma.WeaponWalk(.Heading).FrameCounter = 1
                            .escudo.ShieldWalk(.Heading).FrameCounter = 1
                            .haciendoataque = 0
                        End If
                       
                End With

                    PixelOffsetXTemp = PixelOffsetXTemp + TempChar.MoveOffset.X
                    PixelOffsetYTemp = PixelOffsetYTemp + TempChar.MoveOffset.Y
                    iPPx = ((32 * ScreenX) - 32) + PixelOffsetXTemp + 32
                    iPPy = ((32 * ScreenY) - 32) + PixelOffsetYTemp + 32
                    
                    
                    If Len(TempChar.Nombre) = 0 Then    'NPC

                        If TempChar.Aura_index > 0 Then
                            If Aura(TempChar.Aura_index).Giratoria = 1 Then
                                TempChar.Aura_Angle = TempChar.Aura_Angle + 0.0004
                                If TempChar.Aura_Angle >= 180 Then TempChar.Aura_Angle = 0
                            End If

                            If Aura(TempChar.Aura_index).offset > 0 Then
                                Call Draw_Aura(Aura(TempChar.Aura_index).Aura, iPPx, iPPy + 35 - Aura(TempChar.Aura_index).offset, 48, 52, True, TempChar.Aura_Angle, TempChar.Aura_index)
                            Else
                                Call Draw_Aura(Aura(TempChar.Aura_index).Aura, iPPx, iPPy + 30, 48, 52, True, TempChar.Aura_Angle, TempChar.Aura_index)
                            End If
                        End If
                        'cuerpo npc
                        Call DDrawSombraGrhToSurface(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 0, 1, , X, Y, , 1)
                        Call DDrawTransGrhtoSurface(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value)
                        'cabeza npc
                        If TempChar.Head.Head(TempChar.Heading).GrhIndex > 0 Then
                            Call DDrawSombraGrhToSurface(TempChar.Head.Head(TempChar.Heading), iPPx + 18, iPPy + TempChar.Body.HeadOffset.Y - 17, 1, 0, 1, , X, Y, , 1)
                            Call DDrawTransGrhtoSurface(TempChar.Head.Head(TempChar.Heading), iPPx, iPPy + TempChar.Body.HeadOffset.Y, 1, 1, MapData(X, Y).light_value)
                        End If
                    Else
                        If TempChar.Navegando = 1 Then
                            'Cuerpo (Barca / Galeon / Galera)
                            If Not CharList(MapData(X, Y).CharIndex).invisible And Not UserEstado = 1 Then
                                If TempChar.Aura_index > 0 Then
                                    If Aura(TempChar.Aura_index).Giratoria = 1 Then
                                        TempChar.Aura_Angle = TempChar.Aura_Angle + 0.0004
                                        If TempChar.Aura_Angle >= 180 Then TempChar.Aura_Angle = 0
                                    End If

                                    If Aura(TempChar.Aura_index).offset > 0 Then
                                        Call Draw_Aura(Aura(TempChar.Aura_index).Aura, iPPx, iPPy + 35 - Aura(TempChar.Aura_index).offset, 48, 52, True, TempChar.Aura_Angle, TempChar.Aura_index)
                                    Else
                                        Call Draw_Aura(Aura(TempChar.Aura_index).Aura, iPPx, iPPy + 30, 48, 52, True, TempChar.Aura_Angle, TempChar.Aura_index)
                                    End If
                                End If
                                Call DDrawSombraGrhToSurface(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 0, 1, , X, Y, , 1)
                                Call DDrawTransGrhtoSurface(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value)
                            Else
                                Call DDrawTransGrhtoSurface(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value, , , , True)
                            End If
                        ElseIf Not CharList(MapData(X, Y).CharIndex).invisible And TempChar.Head.Head(TempChar.Heading).GrhIndex > 0 And Not UserEstado = 1 Then

                            If TempChar.Aura_index > 0 Then
                                If Aura(TempChar.Aura_index).Giratoria = 1 Then
                                    TempChar.Aura_Angle = TempChar.Aura_Angle + 0.0004
                                    If TempChar.Aura_Angle >= 180 Then TempChar.Aura_Angle = 0
                                End If

                                If Aura(TempChar.Aura_index).offset > 0 Then
                                    Call Draw_Aura(Aura(TempChar.Aura_index).Aura, iPPx, iPPy + 35 - Aura(TempChar.Aura_index).offset, 48, 52, True, TempChar.Aura_Angle, TempChar.Aura_index)
                                Else
                                    Call Draw_Aura(Aura(TempChar.Aura_index).Aura, iPPx, iPPy + 30, 48, 52, True, TempChar.Aura_Angle, TempChar.Aura_index)
                                End If
                            End If
                            If TempChar.Heading = SOUTH Then
                                If TempChar.Alas.Walk(TempChar.Heading).GrhIndex <> 0 Then
                                    Call DDrawTransGrhtoSurface(TempChar.Alas.Walk(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y + 25, 1, 1, MapData(X, Y).light_value)
                                End If
                            End If
                            'Cuerpo
                            Call DDrawSombraGrhToSurface(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 0, 1, , X, Y, , 1)
                            Call DDrawTransGrhtoSurface(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value)
                            'Cabeza
                            If TempChar.Head.Head(TempChar.Heading).GrhIndex > 0 Then
                                Call DDrawSombraGrhToSurface(TempChar.Head.Head(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X + 19, iPPy + TempChar.Body.HeadOffset.Y - 19, 1, 0, 1, , X, Y, , 1)
                                Call DDrawTransGrhtoSurface(TempChar.Head.Head(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y, 1, 0, MapData(X, Y).light_value)
                            End If
                            If TempChar.casco.Head(TempChar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(TempChar.casco.Head(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y, 1, 0, MapData(X, Y).light_value)
                            If TempChar.arma.WeaponWalk(TempChar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(TempChar.arma.WeaponWalk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value)
                            If TempChar.escudo.ShieldWalk(TempChar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(TempChar.escudo.ShieldWalk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value)
                            If TempChar.Heading <> SOUTH Then
                                If TempChar.Alas.Walk(TempChar.Heading).GrhIndex <> 0 Then
                                    Call DDrawTransGrhtoSurface(TempChar.Alas.Walk(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y + IIf(TempChar.Heading = NORTH, 25, 30), 1, 1, MapData(X, Y).light_value)
                                End If
                            End If
                            
                    
                    
                        ElseIf CharList(MapData(X, Y).CharIndex).invisible And (CharList(MapData(X, Y).CharIndex).Nombre = CharList(UserCharIndex).Nombre Or AmigoClan(MapData(X, Y).CharIndex)) Or CharList(MapData(X, Y).CharIndex).muerto Then
                            Call DDrawTransGrhtoSurface(TempChar.Body.Walk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value, , , , True)
                            If TempChar.Head.Head(TempChar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(TempChar.Head.Head(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y, 1, 1, MapData(X, Y).light_value, , , , True)
                            If TempChar.casco.Head(TempChar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(TempChar.casco.Head(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y, 1, 1, MapData(X, Y).light_value, , , , True)
                            If TempChar.arma.WeaponWalk(TempChar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(TempChar.arma.WeaponWalk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value, , , , True)
                            If TempChar.escudo.ShieldWalk(TempChar.Heading).GrhIndex > 0 Then Call DDrawTransGrhtoSurface(TempChar.escudo.ShieldWalk(TempChar.Heading), iPPx, iPPy, 1, 1, MapData(X, Y).light_value, , , , True)
                            If TempChar.Alas.Walk(TempChar.Heading).GrhIndex <> 0 Then Call DDrawTransGrhtoSurface(TempChar.Alas.Walk(TempChar.Heading), iPPx + TempChar.Body.HeadOffset.X, iPPy + TempChar.Body.HeadOffset.Y + IIf(TempChar.Heading = NORTH, 25, 30), 1, 1, MapData(X, Y).light_value, , , , True)
                        End If

                        If Nombres Then

                            If Not (TempChar.invisible) Then    'visible
                                Dim lCenter As Long
                                If InStr(TempChar.Nombre, "<") > 0 And InStr(TempChar.Nombre, ">") > 0 Then    'con clan
                                    Dim sClan As String

                                    lCenter = (frmMain.textwidth(Left$(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1)) / 2) - 16
                                    sClan = mid$(TempChar.Nombre, InStr(TempChar.Nombre, "<"))

                                    If TempChar.EsPremium Then    'premium
                                        Call DDrawTransGrhtoSurface(estrella, (iPPx - lCenter) - 25, iPPy + 19, 1, 1, MapData(X, Y).light_value)
                                    End If
                                    If Colorinicial <> 139 Then
                                        Call Texto.Text_Draw(iPPx - lCenter, iPPy + 30, Left$(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), D3DColorARGB(Colorinicial, rG(TempChar.Criminal, 1), rG(TempChar.Criminal, 2), rG(TempChar.Criminal, 3)))
                                        lCenter = (frmMain.textwidth(sClan) / 2) - 16
                                        Call Texto.Text_Draw(iPPx - lCenter, iPPy + 45, sClan, D3DColorARGB(Colorinicial, rG(TempChar.Criminal, 1), rG(TempChar.Criminal, 2), rG(TempChar.Criminal, 3)))
                                    Else
                                        lCenter = (frmMain.textwidth(Left$(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1)) / 2) - 16
                                        Call Texto.Text_Draw(iPPx - lCenter, iPPy + 30, Left$(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), D3DColorXRGB(rG(TempChar.Criminal, 1), rG(TempChar.Criminal, 2), rG(TempChar.Criminal, 3)))
                                        lCenter = (frmMain.textwidth(sClan) / 2) - 16
                                        Call Texto.Text_Draw(iPPx - lCenter, iPPy + 45, sClan, D3DColorXRGB(rG(TempChar.Criminal, 1), rG(TempChar.Criminal, 2), rG(TempChar.Criminal, 3)))
                                    End If
                                Else

                                    lCenter = (frmMain.textwidth(TempChar.Nombre) / 2) - 16
                                    If TempChar.EsPremium Then    'premium
                                        Call DDrawTransGrhtoSurface(estrella, (iPPx - lCenter) - 25, iPPy + 19, 1, 1, MapData(X, Y).light_value)
                                    End If
                                    If Colorinicial <> 139 Then
                                        Call Texto.Text_Draw(iPPx - lCenter, iPPy + 30, TempChar.Nombre, D3DColorARGB(Colorinicial, rG(TempChar.Criminal, 1), rG(TempChar.Criminal, 2), rG(TempChar.Criminal, 3)))
                                    Else
                                        Call Texto.Text_Draw(iPPx - lCenter, iPPy + 30, TempChar.Nombre, D3DColorXRGB(rG(TempChar.Criminal, 1), rG(TempChar.Criminal, 2), rG(TempChar.Criminal, 3)))
                                    End If
                                End If

                            ElseIf (TempChar.invisible And MapData(X, Y).CharIndex = UserCharIndex Or (TempChar.Navegando = 1)) Or AmigoClan(MapData(X, Y).CharIndex) Then    'invisible
                                If InStr(TempChar.Nombre, "<") > 0 And InStr(TempChar.Nombre, ">") > 0 Then
                                    lCenter = (frmMain.textwidth(Left$(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1)) / 2) - 16
                                    sClan = mid$(TempChar.Nombre, InStr(TempChar.Nombre, "<"))
                                    If TempChar.EsPremium Then    'premium
                                        Call DDrawTransGrhtoSurface(estrella, (iPPx - lCenter) - 25, iPPy + 19, 1, 1, MapData(X, Y).light_value)
                                    End If
                                    If Colorinicial <> 139 Then
                                        Call Texto.Text_Draw(iPPx - lCenter, iPPy + 30, Left$(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), D3DColorARGB(Colorinicial, rG(TempChar.Criminal, 1), rG(TempChar.Criminal, 2), rG(TempChar.Criminal, 3)))
                                        lCenter = (frmMain.textwidth(sClan) / 2) - 16
                                        Call Texto.Text_Draw(iPPx - lCenter, iPPy + 45, sClan, D3DColorARGB(Colorinicial, rG(TempChar.Criminal, 1), rG(TempChar.Criminal, 2), rG(TempChar.Criminal, 3)))
                                    Else
                                        lCenter = (frmMain.textwidth(Left$(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1)) / 2) - 16
                                        Call Texto.Text_Draw(iPPx - lCenter, iPPy + 30, Left$(TempChar.Nombre, InStr(TempChar.Nombre, "<") - 1), D3DColorXRGB(rG(TempChar.Criminal, 1), rG(TempChar.Criminal, 2), rG(TempChar.Criminal, 3)))
                                        lCenter = (frmMain.textwidth(sClan) / 2) - 16
                                        Call Texto.Text_Draw(iPPx - lCenter, iPPy + 45, sClan, D3DColorXRGB(rG(TempChar.Criminal, 1), rG(TempChar.Criminal, 2), rG(TempChar.Criminal, 3)))
                                    End If
                                Else
                                    lCenter = (frmMain.textwidth(TempChar.Nombre) / 2) - 16
                                    If TempChar.EsPremium Then    'premium
                                        Call DDrawTransGrhtoSurface(estrella, (iPPx - lCenter) - 25, iPPy + 19, 1, 1, MapData(X, Y).light_value)
                                    End If
                                    If Colorinicial <> 139 Then
                                        Call Texto.Text_Draw(iPPx - lCenter, iPPy + 30, TempChar.Nombre, D3DColorARGB(Colorinicial, rG(TempChar.Criminal, 1), rG(TempChar.Criminal, 2), rG(TempChar.Criminal, 3)))
                                    Else
                                        Call Texto.Text_Draw(iPPx - lCenter, iPPy + 30, TempChar.Nombre, D3DColorXRGB(rG(TempChar.Criminal, 1), rG(TempChar.Criminal, 2), rG(TempChar.Criminal, 3)))
                                    End If
                                End If
                            End If
                        End If
                    End If
                    If Dialogos.CantidadDialogos > 0 Then Call Dialogos.Update_Dialog_Pos((iPPx + TempChar.Body.HeadOffset.X), (iPPy + TempChar.Body.HeadOffset.Y), MapData(X, Y).CharIndex)

                    CharList(MapData(X, Y).CharIndex) = TempChar

                    If CharList(MapData(X, Y).CharIndex).FX <> 0 Then Call DDrawTransGrhtoSurface(FxData(TempChar.FX).FX, iPPx + FxData(TempChar.FX).OffSetX, iPPy + FxData(TempChar.FX).OffSetY, 1, 1, MapData(X, Y).light_value, MapData(X, Y).CharIndex, , , True)

                End If
                '*************************************************


                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(3), _
                                                ScreenX * 32 + PixelOffsetX, _
                                                ScreenY * 32 + PixelOffsetY, _
                                                1, 1, .light_value)
                End If
                '************************************************

            End With

            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y
    ScreenY = minYOffset - 5

    ScreenY = minYOffset - TileBufferSize


    For Y = minY To maxY
        ScreenX = minXOffset - TileBufferSize
        For X = minX To maxX
            If MapData(X, Y).particle_group_index And Colorinicial <> 139 Then
                Particle_Group_Render MapData(X, Y).particle_group_index, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, True    '+ (16)
            ElseIf MapData(X, Y).particle_group_index And Colorinicial = 139 Then
                Particle_Group_Render MapData(X, Y).particle_group_index, ScreenX * 32 + PixelOffsetX, ScreenY * 32 + PixelOffsetY, False    '+ (16)
            End If
            ScreenX = ScreenX + 1
        Next X
        ScreenY = ScreenY + 1
    Next Y

    'Particle Layer**********************************
    Call Effect_UpdateAll
    '************************************************

    'Draw blocked tiles and grid
    
    If bTecho Then
        If bTechoAB > 0 Then
            bTechoAB = bTechoAB - 2 'timerTicksPerFrame * 15
            If bTechoAB < 5 Then bTechoAB = 0
        End If
    Else
        If bTechoAB < 255 Then
            bTechoAB = bTechoAB + 2 ' timerTicksPerFrame * 15
            If bTechoAB > 245 Then bTechoAB = 255
        End If
    End If

    If bTechoAB > 0 Then
        ScreenY = minYOffset - TileBufferSize
        For Y = minY To maxY
            ScreenX = minXOffset - TileBufferSize
            For X = minX To maxX
                'Layer 4 **********************************
                If MapData(X, Y).Graphic(4).GrhIndex Then
                    Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), _
                                                ScreenX * 32 + PixelOffsetX, _
                                                ScreenY * 32 + PixelOffsetY, _
                                                1, 1, MapData(X, Y).light_value, , X, Y, , , , bTechoAB)
                End If
                '**********************************

                ScreenX = ScreenX + 1
            Next X
            ScreenY = ScreenY + 1
        Next Y
    End If
    If Niebla Then
        WeatherDoFog = 10
        Engine_Weather_UpdateFog
    End If

    'Engine_Weather_Update

    LastOffsetX = ParticleOffsetX
    LastOffsetY = ParticleOffsetY


End Sub
Public Function RenderSounds()

    If Opciones.sMusica <> CONST_DESHABILITADA Then
        If Opciones.sMusica <> CONST_DESHABILITADA Then
            Sound.NextMusic = CurrentMP3
            Sound.Fading = 350
        End If
    End If
                

End Function

Function InitTileEngine(ByRef setDisplayFormhWnd As Long, setTilePixelHeight As Integer, setTilePixelWidth As Integer, setWindowTileHeight As Integer, setWindowTileWidth As Integer, Optional ByVal CharIndex As Integer) As Boolean
    bTechoAB = 255
    AlphaNPC = 255

    IniPath = App.Path & "\RECURSOS\Init\"

    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder

    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth

    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)

    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock

    Call LoadGrhData
    Call CargarParticulas
    Call CargarAuras
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    Call CargarAnimArmas
    Call CargarAnimEscudos

    HalfWindowTileHeight = WindowTileHeight / 2
    HalfWindowTileWidth = WindowTileWidth / 2


    TileBufferSize = 9
    'Parra: Aca inician las variables globales del Directx8


    '****** INIT DirectX ******
    ' Create the root D3D objects

    Dim DispMode As D3DDISPLAYMODE

    Dim D3DWindow As D3DPRESENT_PARAMETERS

    Set SurfaceDB = New clsSurfaceManDynDX8

    Set dX = New DirectX8
    Set D3D = dX.Direct3DCreate()
    Set D3DX = New D3DX8

    D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    'D3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispModeBK

    With D3DWindow
            .Windowed = True
            .SwapEffect = D3DSWAPEFFECT_COPY
            .BackBufferFormat = DispMode.Format
            .BackBufferWidth = 800
            .BackBufferHeight = 600
            .EnableAutoDepthStencil = 1
            .AutoDepthStencilFormat = D3DFMT_D16
            .hDeviceWindow = frmMain.renderer.hWnd
        End With

    DispMode.Format = D3DFMT_X8R8G8B8


    Set d3ddevice = D3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.renderer.hWnd, D3DCREATE_SOFTWARE_VERTEXPROCESSING, D3DWindow)

    Call SurfaceDB.Init(D3DX, d3ddevice, DEFAULT_MEMORY_TO_USE)
    Texto.Text_Init_Settings
    Texto.Text_Init_Textures
    Engine_Init_ParticleEngine
    d3ddevice.SetVertexShader FVF
    d3ddevice.SetRenderState D3DRS_LIGHTING, False
    d3ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
    d3ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    d3ddevice.SetRenderState D3DRS_ALPHABLENDENABLE, True
    d3ddevice.SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    d3ddevice.SetRenderState D3DRS_POINTSPRITE_ENABLE, 1
    'd3ddevice.SetRenderState D3DRS_POINTSCALE_ENABLE, 0
    engineBaseSpeed = 0.017
    ScrollPixelsPerFrame = 9
    
    With General_Connection_RenderRect
            .Top = 0
            .Left = 0
            .Right = 800
            .Bottom = 600

    End With
     
    With RenderRect
            .Top = 0
            .Left = 0
            .Right = frmMain.renderer.Width
            .Bottom = frmMain.renderer.Height
        End With
        
    InitTileEngine = True

End Function
Private Sub DDrawSombraGrhToSurface(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal center As Byte, ByVal Animate As Byte, Optional ByVal KillAnim As Integer = 0, Optional ByVal Alpha As Boolean, Optional ByVal map_x As Byte = 1, Optional ByVal map_y As Byte = 1, Optional ByVal angle As Single, Optional ByVal shadow As Byte = 0)

On Error Resume Next
    Dim iGrhIndex As Integer


If Animate Then
    If Grh.Started = 1 Then
       
        Grh.FrameCounter = Grh.FrameCounter + ((timerElapsedTime * 0.1) * GrhData(Grh.GrhIndex).NumFrames / Grh.SpeedCounter)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
               
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                   
                If KillAnim <> 0 Then
                    If CharList(KillAnim).FX > 0 Then
                        If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                          CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes <= 0 Then CharList(KillAnim).FX = 0: Exit Sub
                        End If
                    End If
                End If
            End If
    End If
End If

    If Grh.GrhIndex = 0 Then Exit Sub


    iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

    If center Then
        If GrhData(iGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16
        End If
        If GrhData(iGrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32
        End If
    End If

    If map_x Or map_y = 0 Then map_x = 1: map_y = 1

    Dim shadowRgb(3) As Long
    shadowRgb(0) = D3DColorARGB(100, 0, 0, 0)
    shadowRgb(1) = D3DColorARGB(100, 0, 0, 0)
    shadowRgb(2) = D3DColorARGB(100, 0, 0, 0)
    shadowRgb(3) = D3DColorARGB(100, 0, 0, 0)

    Device_Box_Textured_Render_Advance iGrhIndex, _
                                       X, Y, _
                                       GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
                                       shadowRgb(), _
                                       GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
                                       Alpha, angle, shadow

End Sub
Public Sub DeInitTileEngine()

    Dim i As Long
    Dim j As Long

    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                Set .SurfaceEntry(j).Texture = Nothing
            Next j

            'Destroy the arrays
            Erase .SurfaceEntry
        End With
    Next i

    For i = 1 To UBound(ParticleTexture)

        If Not ParticleTexture(i) Is Nothing Then Set ParticleTexture(i) = Nothing

    Next i

    For i = 1 To 4
        If Not SangreTexture(i) Is Nothing Then Set SangreTexture(i) = Nothing
    Next i

    Set dX = Nothing
    Set D3D = Nothing
    Set D3DX = Nothing
    Set d3ddevice = Nothing
    'Set Audio = Nothing
    
    Erase CharList
    Erase Grh
    Erase GrhData
    Erase MapData
    '// Destroy texts
    Call Texto.Text_Destroy
    Sound.Sound_Stop_All
    Sound.Music_Stop
'    Sound = Nothing
    If App.exeName = "Lhirius AO" Then StopURLDetect
'    Detectar 0, 0
End Sub

Public Function ARGB(ByVal R As Long, ByVal G As Long, ByVal b As Long, ByVal a As Long) As Long

    Dim c As Long

    If a > 127 Then
        a = a - 128
        c = a * 2 ^ 24 Or &H80000000
        c = c Or R * 2 ^ 16
        c = c Or G * 2 ^ 8
        c = c Or b
    Else
        c = a * 2 ^ 24
        c = c Or R * 2 ^ 16
        c = c Or G * 2 ^ 8
        c = c Or b
    End If

    ARGB = c

End Function
Private Sub Device_Box_Textured_Render_Advance(ByVal GrhIndex As Long, ByVal dest_x As Integer, ByVal dest_y As Integer, ByVal src_width As Integer, _
                                               ByVal src_height As Integer, ByRef Rgb_List() As Long, ByVal src_x As Integer, _
                                               ByVal src_y As Integer, Optional ByVal alpha_blend As Boolean, Optional ByVal angle As Single, _
                                               Optional ByVal shadow As Byte = 0, _
                                               Optional ByVal Invert_x As Boolean = False, _
                                               Optional ByVal Invert_y As Boolean = False, Optional alphaa As Byte = 255)
    Static src_rect As RECT
    Static dest_rect As RECT
    Static temp_verts(3) As TLVERTEX
    Static d3dTextures As D3D8Textures
    Static light_value(0 To 3) As Long

    If GrhIndex = 0 Then Exit Sub
    Set d3dTextures.Texture = SurfaceDB.GetTexture(GrhData(GrhIndex).FileNum, d3dTextures.texwidth, d3dTextures.TexHeight)

    light_value(0) = Rgb_List(0)
    light_value(1) = Rgb_List(1)
    light_value(2) = Rgb_List(2)
    light_value(3) = Rgb_List(3)

    If (light_value(0) = 0) Then light_value(0) = base_light
    If (light_value(1) = 0) Then light_value(1) = base_light
    If (light_value(2) = 0) Then light_value(2) = base_light
    If (light_value(3) = 0) Then light_value(3) = base_light

    If alphaa <> 255 Then
        Dim aux As D3DCOLORVALUE
        ARGBtoD3DCOLORVALUE RGB(139, 139, 139), aux
        Dim i As Long
        For i = 0 To 3
            light_value(i) = D3DColorARGB(alphaa, aux.R, aux.G, aux.b)
        Next i
    End If


    'Set up the source rectangle
    With src_rect
        .Bottom = src_y + src_height
        .Left = src_x
        .Right = src_x + src_width
        .Top = src_y
    End With

    'Set up the destination rectangle
    With dest_rect
        .Bottom = dest_y + src_height
        .Left = dest_x
        .Right = dest_x + src_width
        .Top = dest_y
    End With

    'Set up the TempVerts(3) vertices
    Geometry_Create_Box temp_verts(), dest_rect, src_rect, light_value(), d3dTextures.texwidth, d3dTextures.TexHeight, angle, Invert_x, Invert_y

    'Set Textures
    d3ddevice.SetTexture 0, d3dTextures.Texture

    If shadow Then
        temp_verts(1).X = temp_verts(1).X + src_width / 2
        temp_verts(1).Y = temp_verts(1).Y - src_height / 2

        temp_verts(3).X = temp_verts(3).X + src_width
        temp_verts(3).Y = temp_verts(3).Y - src_width
    End If

    If alpha_blend Then
        'Set Rendering for alphablending
        d3ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_ONE
        d3ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_ONE
    End If

    'Draw the triangles that make up our square Textures
    d3ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, temp_verts(0), Len(temp_verts(0))

    If alpha_blend Then
        'Set Rendering for colokeying
        d3ddevice.SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        d3ddevice.SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
    End If

End Sub
Private Function Geometry_Create_TLVertex(ByVal X As Single, ByVal Y As Single, ByVal Z As Single, _
                                          ByVal rhw As Single, ByVal Color As Long, ByVal Specular As Long, tu As Single, _
                                          ByVal tv As Single) As TLVERTEX
    Geometry_Create_TLVertex.X = X
    Geometry_Create_TLVertex.Y = Y
    Geometry_Create_TLVertex.Z = Z
    Geometry_Create_TLVertex.rhw = rhw
    Geometry_Create_TLVertex.Color = Color
    Geometry_Create_TLVertex.Specular = Specular
    Geometry_Create_TLVertex.tu = tu
    Geometry_Create_TLVertex.tv = tv
End Function

Private Sub Geometry_Create_Box(ByRef verts() As TLVERTEX, ByRef dest As RECT, ByRef src As RECT, ByRef Rgb_List() As Long, _
                                Optional ByRef Textures_Width As Long, Optional ByRef Textures_Height As Long, Optional ByVal angle As Single, Optional ByVal Invert_x As Boolean = False, Optional ByVal Invert_y As Boolean = False)
'**************************************************************
'Author: Aaron Perkins
'Modified by Juan Martín Sotuyo Dodero
'Last Modify Date: 11/17/2002
'**************************************************************
    Dim x_center As Single
    Dim y_center As Single
    Dim radius As Single
    Dim x_Cor As Single
    Dim y_Cor As Single
    Dim left_point As Single
    Dim right_point As Single
    Dim Temp As Single
    Dim auxr As RECT

    If angle <> 0 Then
        'Center coordinates on screen of the square
        x_center = dest.Left + (dest.Right - dest.Left) / 2
        y_center = dest.Top + (dest.Bottom - dest.Top) / 2

        'Calculate radius
        radius = Sqr((dest.Right - x_center) ^ 2 + (dest.Bottom - y_center) ^ 2)

        'Calculate left and right points
        Temp = (dest.Right - x_center) / radius
        right_point = Atn(Temp / Sqr(-Temp * Temp + 1))
        left_point = PI - right_point
    End If

    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-left_point - angle) * radius
        y_Cor = y_center - Sin(-left_point - angle) * radius
    End If

    auxr = src

    If angle < 0 Then
        src.Left = auxr.Right
        src.Right = auxr.Left
    End If

    If Invert_x Then
        src.Left = auxr.Right
        src.Right = auxr.Left
    End If

    If Invert_y Then
        src.Top = auxr.Bottom
        src.Bottom = auxr.Top
    End If

    '0 - Bottom left vertex
    If Textures_Width And Textures_Height Then
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, Rgb_List(0), 0, src.Left / Textures_Width, (src.Bottom + 1) / Textures_Height)
    Else
        verts(0) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, Rgb_List(0), 0, 0, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Left
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(left_point - angle) * radius
        y_Cor = y_center - Sin(left_point - angle) * radius
    End If

    '1 - Top left vertex
    If Textures_Width And Textures_Height Then
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, Rgb_List(1), 0, src.Left / Textures_Width, src.Top / Textures_Height)
    Else
        verts(1) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, Rgb_List(1), 0, 0, 1)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Bottom
    Else
        x_Cor = x_center + Cos(-right_point - angle) * radius
        y_Cor = y_center - Sin(-right_point - angle) * radius
    End If

    '2 - Bottom right vertex
    If Textures_Width And Textures_Height Then
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, Rgb_List(2), 0, (src.Right + 1) / Textures_Width, (src.Bottom + 1) / Textures_Height)
    Else
        verts(2) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, Rgb_List(2), 0, 1, 0)
    End If
    'Calculate screen coordinates of sprite, and only rotate if necessary
    If angle = 0 Then
        x_Cor = dest.Right
        y_Cor = dest.Top
    Else
        x_Cor = x_center + Cos(right_point - angle) * radius
        y_Cor = y_center - Sin(right_point - angle) * radius
    End If

    '3 - Top right vertex
    If Textures_Width And Textures_Height Then
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, Rgb_List(3), 0, (src.Right + 1) / Textures_Width, src.Top / Textures_Height)
    Else
        verts(3) = Geometry_Create_TLVertex(x_Cor, y_Cor, 0, 1, Rgb_List(3), 0, 1, 1)
    End If

End Sub

'********************************************************
'PARTICULAS ORE 1.0

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''[PARTICULAS]''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function Particle_Group_Create(ByVal map_x As Integer, ByVal map_y As Integer, ByRef grh_index_list() As Long, ByRef Rgb_List() As Long, _
                                      Optional ByVal particle_count As Long = 20, Optional ByVal stream_type As Long = 1, _
                                      Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                      Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                      Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                      Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                      Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                      Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                      Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                      Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                      Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                                      Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                      Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                      Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                      Optional grh_resizex As Integer, Optional grh_resizey As Integer, _
                                      Optional ConLuz As Boolean = True)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/14/2003
'Returns the particle_group_index if successful, else 0
'Modified by Juan Martín Sotuyo Dodero
'Modified by Augusto José Rando
'**************************************************************

    If (map_x <> -1) And (map_y <> -1) Then
        If Map_Particle_Group_Get(map_x, map_y) = 0 Then
            Particle_Group_Create = Particle_Group_Next_Open
            Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), Rgb_List(), alpha_blend, alive_counter, frame_speed, id, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
        End If
    Else
        Particle_Group_Create = Particle_Group_Next_Open
        Particle_Group_Make Particle_Group_Create, map_x, map_y, particle_count, stream_type, grh_index_list(), Rgb_List(), alpha_blend, alive_counter, frame_speed, id, X1, Y1, angle, vecx1, vecx2, vecy1, vecy2, life1, life2, fric, spin_speedL, gravity, grav_strength, bounce_strength, X2, Y2, XMove, move_x1, move_x2, move_y1, move_y2, YMove, spin_speedH, spin, grh_resize, grh_resizex, grh_resizey
    End If

    'If ConLuz = True Then 'Thusing
    'Light.Light_Create map_x, map_y, 5, , 255, 255, 255
    'If YaPrendioLuces = True Then
    '    Light.Light_Render_All
    'End If
    'End If

End Function
Public Function Particle_Group_Remove(ByVal particle_group_index As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
'Make sure it's a legal index
    If Particle_Group_Check(particle_group_index) Then
        Particle_Group_Destroy particle_group_index
        Particle_Group_Remove = True
    End If
End Function

Public Function Particle_Group_Remove_All() As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'*****************************************************************
    Dim Index As Long

    For Index = 1 To particle_group_last
        'Make sure it's a legal index
        If Particle_Group_Check(Index) Then
            Particle_Group_Destroy Index
        End If
    Next Index

    Particle_Group_Remove_All = True
End Function

Public Function Particle_Group_Find(ByVal id As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'Find the index related to the handle
'*****************************************************************
    On Error GoTo ErrorHandler:
    Dim loopc As Long

    loopc = 1
    Do Until particle_group_list(loopc).id = id
        If loopc = particle_group_last Then
            Particle_Group_Find = 0
            Exit Function
        End If
        loopc = loopc + 1
    Loop

    Particle_Group_Find = loopc
    Exit Function
ErrorHandler:
    Particle_Group_Find = 0
End Function

Private Sub Particle_Group_Destroy(ByVal particle_group_index As Long)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'**************************************************************
    Dim Temp As particle_group

    If particle_group_list(particle_group_index).map_x > 0 And particle_group_list(particle_group_index).map_y > 0 Then
        MapData(particle_group_list(particle_group_index).map_x, particle_group_list(particle_group_index).map_y).particle_group_index = 0
    ElseIf particle_group_list(particle_group_index).char_index Then
        If Char_Check(particle_group_list(particle_group_index).char_index) Then
            'For I = 1 To charlist(particle_group_list(particle_group_index).char_index).particle_count
            '    If charlist(particle_group_list(particle_group_index).char_index).particle_group(I) = particle_group_index Then
            '        charlist(particle_group_list(particle_group_index).char_index).particle_group(I) = 0
            '
            '        Exit For
            '    End If
            'Next I
        End If
    End If

    particle_group_list(particle_group_index) = Temp

    'Update array size
    If particle_group_index = particle_group_last Then
        Do Until particle_group_list(particle_group_last).active
            particle_group_last = particle_group_last - 1
            If particle_group_last = 0 Then
                particle_group_count = 0
                Exit Sub
            End If
        Loop
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count - 1
End Sub

Private Sub Particle_Group_Make(ByVal particle_group_index As Long, ByVal map_x As Integer, ByVal map_y As Integer, _
                                ByVal particle_count As Long, ByVal stream_type As Long, ByRef grh_index_list() As Long, ByRef Rgb_List() As Long, _
                                Optional ByVal alpha_blend As Boolean, Optional ByVal alive_counter As Long = -1, _
                                Optional ByVal frame_speed As Single = 0.5, Optional ByVal id As Long, _
                                Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                                Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                                Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                                Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                                Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                                Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                                Optional bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                                Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                                Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                                Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                                Optional grh_resizex As Integer, Optional grh_resizey As Integer)

'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Last Modify Date: 5/15/2003
'Makes a new particle effect
'Modified by Juan Martín Sotuyo Dodero
'*****************************************************************
'Update array size
    If particle_group_index > particle_group_last Then
        particle_group_last = particle_group_index
        ReDim Preserve particle_group_list(1 To particle_group_last)
    End If
    particle_group_count = particle_group_count + 1

    'Make active
    particle_group_list(particle_group_index).active = True

    'Map pos
    If (map_x <> -1) And (map_y <> -1) Then
        particle_group_list(particle_group_index).map_x = map_x
        particle_group_list(particle_group_index).map_y = map_y
    End If

    'Grh list
    ReDim particle_group_list(particle_group_index).grh_index_list(1 To UBound(grh_index_list))
    particle_group_list(particle_group_index).grh_index_list() = grh_index_list()
    particle_group_list(particle_group_index).grh_index_count = UBound(grh_index_list)

    'Sets alive vars
    If alive_counter = -1 Then
        particle_group_list(particle_group_index).alive_counter = -1
        particle_group_list(particle_group_index).never_die = True
    Else
        particle_group_list(particle_group_index).alive_counter = alive_counter
        particle_group_list(particle_group_index).never_die = False
    End If

    'alpha blending
    particle_group_list(particle_group_index).alpha_blend = alpha_blend

    'stream type
    particle_group_list(particle_group_index).stream_type = stream_type

    'speed
    particle_group_list(particle_group_index).frame_speed = frame_speed

    particle_group_list(particle_group_index).X1 = X1
    particle_group_list(particle_group_index).Y1 = Y1
    particle_group_list(particle_group_index).X2 = X2
    particle_group_list(particle_group_index).Y2 = Y2
    particle_group_list(particle_group_index).angle = angle
    particle_group_list(particle_group_index).vecx1 = vecx1
    particle_group_list(particle_group_index).vecx2 = vecx2
    particle_group_list(particle_group_index).vecy1 = vecy1
    particle_group_list(particle_group_index).vecy2 = vecy2
    particle_group_list(particle_group_index).life1 = life1
    particle_group_list(particle_group_index).life2 = life2
    particle_group_list(particle_group_index).fric = fric
    particle_group_list(particle_group_index).spin = spin
    particle_group_list(particle_group_index).spin_speedL = spin_speedL
    particle_group_list(particle_group_index).spin_speedH = spin_speedH
    particle_group_list(particle_group_index).gravity = gravity
    particle_group_list(particle_group_index).grav_strength = grav_strength
    particle_group_list(particle_group_index).bounce_strength = bounce_strength
    particle_group_list(particle_group_index).XMove = XMove
    particle_group_list(particle_group_index).YMove = YMove
    particle_group_list(particle_group_index).move_x1 = move_x1
    particle_group_list(particle_group_index).move_x2 = move_x2
    particle_group_list(particle_group_index).move_y1 = move_y1
    particle_group_list(particle_group_index).move_y2 = move_y2

    particle_group_list(particle_group_index).Rgb_List(0) = Rgb_List(0)
    particle_group_list(particle_group_index).Rgb_List(1) = Rgb_List(1)
    particle_group_list(particle_group_index).Rgb_List(2) = Rgb_List(2)
    particle_group_list(particle_group_index).Rgb_List(3) = Rgb_List(3)

    particle_group_list(particle_group_index).grh_resize = grh_resize
    particle_group_list(particle_group_index).grh_resizex = grh_resizex
    particle_group_list(particle_group_index).grh_resizey = grh_resizey

    'handle
    particle_group_list(particle_group_index).id = id

    'create particle stream
    particle_group_list(particle_group_index).particle_count = particle_count
    ReDim particle_group_list(particle_group_index).particle_stream(1 To particle_count)

    'plot particle group on map
    If (map_x <> -1) And (map_y <> -1) Then
        MapData(map_x, map_y).particle_group_index = particle_group_index
    End If

End Sub
Public Function Particle_Type_Get(ByVal particle_index As Long) As Long
'*****************************************************************
'Author: Juan Martín Sotuyo Dodero (juansotuyo@hotmail.com)
'Last Modify Date: 8/27/2003
'Returns the stream type of a particle stream
'*****************************************************************
    If Particle_Group_Check(particle_index) Then
        Particle_Type_Get = particle_group_list(particle_index).stream_type
    End If
End Function
Private Sub Particle_Group_Render(ByVal particle_group_index As Long, ByVal screen_x As Integer, ByVal screen_y As Integer, Optional UsaColor As Boolean = False)
'*****************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'Renders a particle stream at a paticular screen point
'*****************************************************************
    Dim loopc As Long
    Dim temp_rgb(0 To 3) As Long
    Dim no_move As Boolean

    'Set colors
    If Not UsaColor Then
        temp_rgb(0) = particle_group_list(particle_group_index).Rgb_List(0)
        temp_rgb(1) = particle_group_list(particle_group_index).Rgb_List(1)
        temp_rgb(2) = particle_group_list(particle_group_index).Rgb_List(2)
        temp_rgb(3) = particle_group_list(particle_group_index).Rgb_List(3)
    Else
        temp_rgb(0) = ARGB(Colorinicial, Colorinicial, Colorinicial, 255)
        temp_rgb(1) = ARGB(Colorinicial, Colorinicial, Colorinicial, 255)
        temp_rgb(2) = ARGB(Colorinicial, Colorinicial, Colorinicial, 255)
        temp_rgb(3) = ARGB(Colorinicial, Colorinicial, Colorinicial, 255)
    End If

    If particle_group_list(particle_group_index).alive_counter Then

        'See if it is time to move a particle
        particle_group_list(particle_group_index).frame_counter = particle_group_list(particle_group_index).frame_counter + timerTicksPerFrame
        If particle_group_list(particle_group_index).frame_counter > particle_group_list(particle_group_index).frame_speed Then
            particle_group_list(particle_group_index).frame_counter = 0
            no_move = False
        Else
            no_move = True
        End If



        'If it's still alive render all the particles inside
        For loopc = 1 To particle_group_list(particle_group_index).particle_count

            'Render particle
            Particle_Render particle_group_list(particle_group_index).particle_stream(loopc), _
                            screen_x, screen_y, _
                            particle_group_list(particle_group_index).grh_index_list(Round(RandomNumber(1, particle_group_list(particle_group_index).grh_index_count), 0)), _
                            temp_rgb(), _
                            particle_group_list(particle_group_index).alpha_blend, no_move, _
                            particle_group_list(particle_group_index).X1, particle_group_list(particle_group_index).Y1, particle_group_list(particle_group_index).angle, _
                            particle_group_list(particle_group_index).vecx1, particle_group_list(particle_group_index).vecx2, _
                            particle_group_list(particle_group_index).vecy1, particle_group_list(particle_group_index).vecy2, _
                            particle_group_list(particle_group_index).life1, particle_group_list(particle_group_index).life2, _
                            particle_group_list(particle_group_index).fric, particle_group_list(particle_group_index).spin_speedL, _
                            particle_group_list(particle_group_index).gravity, particle_group_list(particle_group_index).grav_strength, _
                            particle_group_list(particle_group_index).bounce_strength, particle_group_list(particle_group_index).X2, _
                            particle_group_list(particle_group_index).Y2, particle_group_list(particle_group_index).XMove, _
                            particle_group_list(particle_group_index).move_x1, particle_group_list(particle_group_index).move_x2, _
                            particle_group_list(particle_group_index).move_y1, particle_group_list(particle_group_index).move_y2, _
                            particle_group_list(particle_group_index).YMove, particle_group_list(particle_group_index).spin_speedH, _
                            particle_group_list(particle_group_index).spin, particle_group_list(particle_group_index).grh_resize, particle_group_list(particle_group_index).grh_resizex, particle_group_list(particle_group_index).grh_resizey
        Next loopc

        If no_move = False Then
            'Update the group alive counter
            If particle_group_list(particle_group_index).never_die = False Then
                particle_group_list(particle_group_index).alive_counter = particle_group_list(particle_group_index).alive_counter - 1
            End If
        End If

    Else
        'If it's dead destroy it
        particle_group_list(particle_group_index).particle_count = particle_group_list(particle_group_index).particle_count - 1
        If particle_group_list(particle_group_index).particle_count <= 0 Then Particle_Group_Destroy particle_group_index
    End If
End Sub

Private Sub Particle_Render(ByRef temp_particle As Particle, ByVal screen_x As Integer, ByVal screen_y As Integer, _
                            ByVal grh_index As Long, ByRef Rgb_List() As Long, _
                            Optional ByVal alpha_blend As Boolean, Optional ByVal no_move As Boolean, _
                            Optional ByVal X1 As Integer, Optional ByVal Y1 As Integer, Optional ByVal angle As Integer, _
                            Optional ByVal vecx1 As Integer, Optional ByVal vecx2 As Integer, _
                            Optional ByVal vecy1 As Integer, Optional ByVal vecy2 As Integer, _
                            Optional ByVal life1 As Integer, Optional ByVal life2 As Integer, _
                            Optional ByVal fric As Integer, Optional ByVal spin_speedL As Single, _
                            Optional ByVal gravity As Boolean, Optional grav_strength As Long, _
                            Optional ByVal bounce_strength As Long, Optional ByVal X2 As Integer, Optional ByVal Y2 As Integer, _
                            Optional ByVal XMove As Boolean, Optional ByVal move_x1 As Integer, Optional ByVal move_x2 As Integer, _
                            Optional ByVal move_y1 As Integer, Optional ByVal move_y2 As Integer, Optional ByVal YMove As Boolean, _
                            Optional ByVal spin_speedH As Single, Optional ByVal spin As Boolean, Optional grh_resize As Boolean, _
                            Optional MoveX As Integer, Optional MoveY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Modified by: Ryan Cain (Onezero)
'Modified by: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/15/2003
'**************************************************************
'On Error GoTo A:


    If no_move = False Then

        If temp_particle.alive_counter = 0 Then
            'Start new particle
            InitGrh temp_particle.Grh, grh_index

            If MoveX <> 0 Or MoveY <> 0 Then
                temp_particle.X = RandomNumber(X1, X2) - (base_tile_size / 2) + screen_x
                temp_particle.Y = RandomNumber(Y1, Y2) - (base_tile_size / 2) + screen_y
            Else

                temp_particle.X = RandomNumber(X1, X2) - (base_tile_size / 2)
                temp_particle.Y = RandomNumber(Y1, Y2) - (base_tile_size / 2)

            End If

            temp_particle.vector_x = RandomNumber(vecx1, vecx2)
            temp_particle.vector_y = RandomNumber(vecy1, vecy2)
            temp_particle.angle = angle
            temp_particle.alive_counter = RandomNumber(life1, life2)
            temp_particle.friction = fric
            temp_particle.Alpha = 255
            temp_particle.TimeAlpha = temp_particle.alive_counter * 0.5
        Else
            'Continue old particle
            'Do gravity
            If gravity = True Then
                temp_particle.vector_y = temp_particle.vector_y + grav_strength
                If temp_particle.Y > 0 Then
                    'bounce
                    temp_particle.vector_y = bounce_strength
                End If
            End If
            'Do rotation
            If spin = True Then temp_particle.angle = temp_particle.angle + (RandomNumber(spin_speedL, spin_speedH) / 100)

            If temp_particle.angle >= 360 Then
                temp_particle.angle = 0
            End If

            If XMove = True Then temp_particle.vector_x = RandomNumber(move_x1, move_x2)
            If YMove = True Then temp_particle.vector_y = RandomNumber(move_y1, move_y2)
        End If

        'Add in vector
        temp_particle.X = temp_particle.X + (temp_particle.vector_x \ temp_particle.friction)
        temp_particle.Y = temp_particle.Y + (temp_particle.vector_y \ temp_particle.friction)

        'decrement counter
        temp_particle.alive_counter = temp_particle.alive_counter - 1
    End If

    'Draw it
    If grh_resize = True Then
        If temp_particle.Grh.GrhIndex Then
            ' Grh_Render_Advance temp_particle.grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, grh_resizex, grh_resizey, rgb_list(),True, True, alpha_blend
            'DDrawTransGrhtoSurface temp_particle.Grh, temp_particle.X, temp_particle.Y, 1, 1, , , , , alpha_blend, , , temp_particle.angle, D3DColorARGB(temp_particle.Alpha, R, G, b)
            Draw_Grh temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, 1, 1, Rgb_List(), alpha_blend
            Exit Sub
        End If
    End If

    If temp_particle.Grh.GrhIndex Then
        'Draw_Grh temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, True, True, rgb_list(), alpha_blend, , , temp_particle.angle
        If (temp_particle.Alpha > 0) And (temp_particle.alive_counter <= temp_particle.TimeAlpha) Then

            temp_particle.Alpha = temp_particle.Alpha - timerTicksPerFrame * 15

        End If
        Draw_Grh temp_particle.Grh, temp_particle.X + screen_x, temp_particle.Y + screen_y, 1, 1, Rgb_List(), alpha_blend

        'Grh_Render temp_particle.Grh, temp_particle.x + screen_x, temp_particle.y + screen_y, rgb_list(),  True, True, alpha_blend
    End If

a:

End Sub

Private Function Particle_Group_Next_Open() As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'
'*****************************************************************
    On Error GoTo ErrorHandler:
    Dim loopc As Long

    loopc = 1
    Do Until particle_group_list(loopc).active = False
        If loopc = particle_group_last Then
            Particle_Group_Next_Open = particle_group_last + 1
            Exit Function
        End If
        loopc = loopc + 1
    Loop

    Particle_Group_Next_Open = loopc
    Exit Function
ErrorHandler:
    Particle_Group_Next_Open = 1
End Function

Private Function Particle_Group_Check(ByVal particle_group_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
'check index
    If particle_group_index > 0 And particle_group_index <= particle_group_last Then
        If particle_group_list(particle_group_index).active Then
            Particle_Group_Check = True
        End If
    End If
End Function
Rem Mannakia .. Parituclas ORE 1.0.

Public Function General_Field_Read(ByVal field_pos As Long, ByVal Text As String, ByVal delimiter As Byte) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets a field from a delimited string
'*****************************************************************
    Dim i As Long
    Dim LastPos As Long
    Dim FieldNum As Long

    LastPos = 0
    FieldNum = 0
    For i = 1 To Len(Text)
        If delimiter = CByte(Asc(mid$(Text, i, 1))) Then
            FieldNum = FieldNum + 1
            If FieldNum = field_pos Then
                General_Field_Read = mid$(Text, LastPos + 1, (InStr(LastPos + 1, Text, Chr$(delimiter), vbTextCompare) - 1) - (LastPos))
                Exit Function
            End If
            LastPos = i
        End If
    Next i
    FieldNum = FieldNum + 1
    If FieldNum = field_pos Then
        General_Field_Read = mid$(Text, LastPos + 1)
    End If
End Function
Public Function General_Var_Get(ByVal File As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Get a var to from a text file
'*****************************************************************
    Dim sSpaces As String    'Input that the program will retrieve
    Dim szReturn As String    'Default value if the string is not found

    szReturn = ""

    sSpaces = Space$(5000)

    getprivateprofilestring Main, Var, szReturn, sSpaces, Len(sSpaces), File

    General_Var_Get = RTrim$(sSpaces)
    General_Var_Get = Left$(General_Var_Get, Len(General_Var_Get) - 1)
End Function
Public Function Map_Particle_Group_Get(ByVal map_x As Long, ByVal map_y As Long) As Long
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 2/20/2003
'Checks to see if a tile position has a particle_group_index and return it
'*****************************************************************
    If Map_In_Bounds(map_x, map_y) Then
        Map_Particle_Group_Get = map_current.map_grid(map_x, map_y).particle_group_index
    Else
        Map_Particle_Group_Get = 0
    End If
End Function
Private Function Char_Check(ByVal char_index As Long) As Boolean
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 1/04/2003
'
'**************************************************************
'check char_index
    If char_index > 0 Then
        If CharList(char_index).active Then
            Char_Check = True
        End If
    End If
End Function
Public Function Map_In_Bounds(ByVal map_x As Long, ByVal map_y As Long) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If map_x < map_current.map_x_min Or map_x > map_current.map_x_max Or map_y < map_current.map_y_min Or map_y > map_current.map_y_max Then
        Map_In_Bounds = False
        Exit Function
    End If

    Map_In_Bounds = True
End Function
'********************************************************
'PARTICULAS ORE 1.0
Sub Draw_Grh(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, ByRef Color() As Long, Optional Alpha As Boolean, Optional ByVal KillAnim As Integer = 0, Optional ByVal map_x As Byte, Optional ByVal map_y As Byte)
'***************************
'/////By Thusing/////
'***************************

On Error Resume Next
    Dim iGrhIndex As Integer


If Animate Then
    If Grh.Started = 1 Then
       
        Grh.FrameCounter = Grh.FrameCounter + ((timerElapsedTime * 0.1) * GrhData(Grh.GrhIndex).NumFrames / Grh.SpeedCounter)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
               
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                   
                If KillAnim <> 0 Then
                    If CharList(KillAnim).FX > 0 Then
                        If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                          CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes <= 0 Then CharList(KillAnim).FX = 0: Exit Sub
                        End If
                    End If
                End If
            End If
    End If
End If

    If Grh.GrhIndex = 0 Then Exit Sub


    iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

    If center Then
        If GrhData(iGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16
        End If
        If GrhData(iGrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32
        End If
    End If

    If map_x Or map_y = 0 Then map_x = 1: map_y = 1

    Device_Box_Textured_Render_Advance iGrhIndex, _
                                       X, Y, _
                                       GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
                                       Color(), _
                                       GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
                                       Alpha
    ' 0, 0, Invert_x, Invert_y

End Sub

Sub Engine_Weather_UpdateFog()
'*****************************************************************
'Update the fog effects
'*****************************************************************
    Dim TempGrh As Grh
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim ElapsedTime As Single
    ElapsedTime = Engine_ElapsedTime

    If WeatherFogCount = 0 Then WeatherFogCount = 13

    WeatherFogX1 = WeatherFogX1 + (ElapsedTime * (0.018 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY1 = WeatherFogY1 + (ElapsedTime * (0.013 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)

    Do While WeatherFogX1 < -512
        WeatherFogX1 = WeatherFogX1 + 512
    Loop
    Do While WeatherFogY1 < -512
        WeatherFogY1 = WeatherFogY1 + 512
    Loop
    Do While WeatherFogX1 > 0
        WeatherFogX1 = WeatherFogX1 - 512
    Loop
    Do While WeatherFogY1 > 0
        WeatherFogY1 = WeatherFogY1 - 512
    Loop

    WeatherFogX2 = WeatherFogX2 - (ElapsedTime * (0.037 + Rnd * 0.01)) + (LastOffsetX - ParticleOffsetX)
    WeatherFogY2 = WeatherFogY2 - (ElapsedTime * (0.021 + Rnd * 0.01)) + (LastOffsetY - ParticleOffsetY)
    Do While WeatherFogX2 < -512
        WeatherFogX2 = WeatherFogX2 + 512
    Loop
    Do While WeatherFogY2 < -512
        WeatherFogY2 = WeatherFogY2 + 512
    Loop
    Do While WeatherFogX2 > 0
        WeatherFogX2 = WeatherFogX2 - 512
    Loop
    Do While WeatherFogY2 > 0
        WeatherFogY2 = WeatherFogY2 - 512
    Loop

    TempGrh.FrameCounter = 1

    'Render fog 2
    TempGrh.GrhIndex = 20600
    X = 2
    Y = -1

    For i = 1 To WeatherFogCount
        Draw_Niebla TempGrh, (X * 512) + WeatherFogX2, (Y * 512) + WeatherFogY2, 1, 1
        X = X + 1
        If X > (1 + (ScreenWidth \ 512)) Then
            X = 0
            Y = Y + 1
        End If
    Next i

    'Render fog 1
    TempGrh.GrhIndex = 20601
    X = 0
    Y = 0

    Dim ColorFog(0) As Long
    ColorFog(0) = D3DColorXRGB(255, 255, 255)

    For i = 1 To WeatherFogCount
        DDrawTransGrhtoSurface TempGrh, (X * 512) + WeatherFogX1, (Y * 512) + WeatherFogY1, 1, 1, ColorFog, , , , , , , Opciones.bGraphics
        X = X + 1
        If X > (2 + (ScreenWidth \ 512)) Then
            X = 0
            Y = Y + 1
        End If
    Next i

End Sub

Function Engine_PixelPosX(ByVal X As Integer) As Integer
    Engine_PixelPosX = (X - 1) * 32
End Function

Function Engine_PixelPosY(ByVal Y As Integer) As Integer
    Engine_PixelPosY = (Y - 1) * 32
End Function

Private Function Engine_ElapsedTime() As Long
    Dim start_time As Long
    start_time = GetTickCount
    Engine_ElapsedTime = start_time - EndTime
    If Engine_ElapsedTime > 1000 Then Engine_ElapsedTime = 1000
    EndTime = start_time
End Function
Sub Draw_Niebla(Grh As Grh, ByVal X As Integer, ByVal Y As Integer, center As Byte, Animate As Byte, Optional ByVal KillAnim As Integer = 0, Optional ByVal map_x As Byte, Optional ByVal map_y As Byte, Optional ByVal angle As Single)

On Error Resume Next
    Dim iGrhIndex As Integer


If Animate Then
    If Grh.Started = 1 Then
       
        Grh.FrameCounter = Grh.FrameCounter + ((timerElapsedTime * 0.1) * GrhData(Grh.GrhIndex).NumFrames / Grh.SpeedCounter)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
               
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                   
                If KillAnim <> 0 Then
                    If CharList(KillAnim).FX > 0 Then
                        If CharList(KillAnim).FxLoopTimes <> LoopAdEternum Then
                          CharList(KillAnim).FxLoopTimes = CharList(KillAnim).FxLoopTimes - 1
                            If CharList(KillAnim).FxLoopTimes <= 0 Then CharList(KillAnim).FX = 0: Exit Sub
                        End If
                    End If
                End If
            End If
    End If
End If

    If Grh.GrhIndex = 0 Then Exit Sub


    iGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)

    If center Then
        If GrhData(iGrhIndex).TileWidth <> 1 Then
            X = X - Int(GrhData(iGrhIndex).TileWidth * 16) + 16
        End If
        If GrhData(iGrhIndex).TileHeight <> 1 Then
            Y = Y - Int(GrhData(iGrhIndex).TileHeight * 32) + 32
        End If
    End If

    If map_x Or map_y = 0 Then map_x = 1: map_y = 1

    Dim cc(3) As Long
    cc(1) = D3DColorARGB(40, 255, 255, 255)
    cc(2) = D3DColorARGB(40, 255, 255, 255)
    cc(3) = D3DColorARGB(40, 255, 255, 255)
    cc(0) = D3DColorARGB(40, 255, 255, 255)

    Device_Box_Textured_Render_Advance iGrhIndex, _
                                       X, Y, _
                                       GrhData(iGrhIndex).pixelWidth, GrhData(iGrhIndex).pixelHeight, _
                                       cc(), _
                                       GrhData(iGrhIndex).sX, GrhData(iGrhIndex).sY, _
                                       , angle

End Sub


Public Function ARGBtoD3DCOLORVALUE(ByVal ARGB As Long, ByRef Color As D3DCOLORVALUE)
'*********************************************************
'****** Coded by Dunkan (emanuel.m@dunkancorp.com) *******
'*********************************************************
    Dim dest(3) As Byte
    CopyMemory dest(0), ARGB, 4
    Color.a = dest(3)
    Color.R = dest(2)
    Color.G = dest(1)
    Color.b = dest(0)
End Function
Sub ActualizarBarras()

    On Error Resume Next

    'energia
    If Energiafalsa <> UserMinSTA Or frmMain.cantidadsta.Caption = "" Then
        If Energiafalsa < UserMinSTA Then
            Energiafalsa = Energiafalsa + (UserMaxSTA / 60)
            If Energiafalsa > UserMinSTA Then Energiafalsa = UserMinSTA
        Else
            Energiafalsa = Energiafalsa - (UserMaxSTA / 60)
            If Energiafalsa < UserMinSTA Then Energiafalsa = UserMinSTA
        End If

        frmMain.STAShp.Width = (((Energiafalsa / 100) / (UserMaxSTA / 100)) * 93)
        frmMain.cantidadsta.Caption = Val(Energiafalsa) & "/" & UserMaxSTA
    End If

    'mana
    If Manafalsa <> UserMinMAN Or frmMain.cantidadmana.Caption = "" Then
        If Manafalsa < UserMinMAN Then
            Manafalsa = Manafalsa + (UserMaxMAN / 60)
            If Manafalsa > UserMinMAN Then Manafalsa = UserMinMAN
        Else
            Manafalsa = Manafalsa - (UserMaxMAN / 60)
            If Manafalsa < UserMinMAN Then Manafalsa = UserMinMAN
        End If

        frmMain.cantidadmana.Caption = Val(Manafalsa) & "/" & UserMaxMAN
        frmMain.MANShp.Width = (((Manafalsa / 100) / (UserMaxMAN / 100)) * 93)
    End If

    'vida
    If Vidafalsa <> UserMinHP Or frmMain.cantidadhp.Caption = "" Then
        If Vidafalsa < UserMinHP Then
            Vidafalsa = Vidafalsa + (UserMaxHP / 60)
            If Vidafalsa > UserMinHP Then Vidafalsa = UserMinHP
        Else
            Vidafalsa = Vidafalsa - (UserMaxHP / 60)
            If Vidafalsa < UserMinHP Then Vidafalsa = UserMinHP
        End If

        frmMain.cantidadhp.Caption = Val(Vidafalsa) & "/" & UserMaxHP
        frmMain.Hpshp.Width = (((Vidafalsa / 100) / (UserMaxHP / 100)) * 93)
    End If

    'comida
    If HambreFalsa <> UserMinHAM Then
        If HambreFalsa < UserMinHAM Then
            HambreFalsa = HambreFalsa + (UserMaxHAM / 60)
            If HambreFalsa > UserMaxHAM Then HambreFalsa = UserMaxHAM
        Else
            HambreFalsa = HambreFalsa - (UserMaxHAM / 60)
            If HambreFalsa < UserMaxHAM Then HambreFalsa = UserMaxHAM
        End If

        frmMain.cantidadhambre.Caption = HambreFalsa & "/" & UserMaxHAM
        frmMain.COMIDAsp.Width = (((HambreFalsa / 100) / (UserMaxHAM / 100)) * 93)
    End If

    'bebida
    If Aguafalsa <> UserMinAGU Then
        If Aguafalsa < UserMinAGU Then
            Aguafalsa = Aguafalsa + (UserMaxAGU / 60)
            If Aguafalsa > UserMaxAGU Then Aguafalsa = UserMaxAGU
        Else
            Aguafalsa = Aguafalsa - (UserMaxAGU / 60)
            If Aguafalsa < UserMaxAGU Then Aguafalsa = UserMaxAGU
        End If

        frmMain.cantidadagua.Caption = Aguafalsa & "/" & UserMaxAGU
        frmMain.AGUAsp.Width = (((Aguafalsa / 100) / (UserMaxAGU / 100)) * 93)
    End If

    If OroFalso <> UserGLD Then
        If OroFalso < UserGLD Then
            OroFalso = OroFalso + (UserGLD / 60)
            If OroFalso > UserGLD Then OroFalso = UserGLD
        Else
            OroFalso = OroFalso - (UserGLD / 60)
            If OroFalso < UserGLD Then OroFalso = UserGLD
        End If
    End If

    frmMain.GldLbl = PonerPuntos(OroFalso)

End Sub


Public Function Map_Item_Grh_In_Current_Area(ByVal grh_index As Long, ByRef x_pos As Integer, ByRef y_pos As Integer) As Boolean
'*****************************************************************
'Author: Augusto José Rando
'Co-Author: Lorwik
'*****************************************************************
    On Error GoTo ErrorHandler
 
    Dim map_x As Integer
    Dim map_y As Integer
    Dim X As Integer, Y As Integer
 
    Call Char_Pos_Get(UserCharIndex, map_x, map_y)
 
    If Map_In_Bounds(map_x, map_y) Then
        For Y = map_y - MinYBorder + 1 To map_y + MinYBorder - 1
          For X = map_x - MinXBorder + 1 To map_x + MinXBorder - 1
                If Y < 1 Then Y = 1
                If X < 1 Then X = 1
                If MapData(X, Y).ObjGrh.GrhIndex = grh_index Then
                    x_pos = X
                    y_pos = Y
                    Map_Item_Grh_In_Current_Area = True
                    Exit Function
                End If
          Next X
        Next Y
    End If
 
    Exit Function
 
ErrorHandler:
    Map_Item_Grh_In_Current_Area = False
 
End Function
 
Public Function Char_Pos_Get(ByVal char_index As Integer, ByRef map_x As Integer, ByRef map_y As Integer) As Boolean
'*****************************************************************
'Author: Aaron Perkins
'Co-Author: Lorwik
'*****************************************************************
   'Make sure it's a legal char_index
    If Char_Check(char_index) Then
        map_x = CharList(char_index).Pos.X
        map_y = CharList(char_index).Pos.Y
        Char_Pos_Get = True
    End If
End Function


 
 Public Sub RenderConnect()
     
    Dim X As Long, Y As Long
     
    Dim Rgb_List(3) As Long
     
    Rgb_List(0) = D3DColorXRGB(255, 255, 255)
    Rgb_List(1) = D3DColorXRGB(255, 255, 255)
    Rgb_List(2) = D3DColorXRGB(255, 255, 255)
    Rgb_List(3) = D3DColorXRGB(255, 255, 255)
     
   
   
 
    d3ddevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 0#, 0
    d3ddevice.BeginScene
       
       
            For X = 1 To 25
                For Y = 1 To 19
                    With MapData(40 + X, 30 + Y)
                        If .Graphic(1).GrhIndex <> 0 Then _
                            DDrawTransGrhtoSurface .Graphic(1), (X - 1) * 32, (Y - 1) * 32, 1, 0, Rgb_List()
                        If .Graphic(2).GrhIndex <> 0 Then _
                            DDrawTransGrhtoSurface .Graphic(2), (X - 1) * 32, (Y - 1) * 32, 1, 1, Rgb_List()
                    End With
                Next Y
            Next X
            
            
            
            
            For X = 1 To 25
                For Y = 1 To 19
                    With MapData(40 + X, 30 + Y)
                        If .Graphic(3).GrhIndex <> 0 Then
                            DDrawTransGrhtoSurface .Graphic(3), (X - 1) * 32, (Y - 1) * 32, 1, 0, Rgb_List()
                        End If
                    End With
                Next Y
            Next X
            
            For X = 1 To 25
                For Y = 1 To 19
                    With MapData(40 + X, 30 + Y)
                        If .ObjGrh.GrhIndex <> 0 Then
                            Call DDrawTransGrhtoSurface(.ObjGrh, _
                                (X - 1) * 32, (Y - 1) * 32, 1, 1, Rgb_List())
                        End If
                    End With
                Next Y
            Next X
            
            For Y = 1 To 25
                For X = 1 To 19
                    With MapData(40 + X, 30 + Y)
                        If .particle_group_index Then
                            Particle_Group_Render .particle_group_index, (X - 1) * 32, (Y - 1) * 32, False   '+ (16)
                        End If
                    End With
                Next X
            Next Y
    
    
            
            For X = 1 To 25
                For Y = 1 To 19
                    With MapData(40 + X, 30 + Y)
                        If .Graphic(4).GrhIndex <> 0 Then
                            DDrawTransGrhtoSurface .Graphic(4), (X - 1) * 32, (Y - 1) * 32, 1, 1, Rgb_List()
                        End If
                    End With
                Next Y
            Next X
           
        Call Connecting_Effect
   
   
d3ddevice.Present General_Connection_RenderRect, ByVal 0, frmConnect.RenderConnect.hWnd, ByVal 0
d3ddevice.EndScene

lFrameLimiter = GetTickCount
FramesPerSecCounter = FramesPerSecCounter + 1
timerElapsedTime = GetElapsedTime()
timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
                
If GetTickCount - lFrameTimer > 1000 Then
    FramesPerSec = FramesPerSecCounter
    'If FramesPerSec <> 0 Then ScrollPixelsPerFrame = 150 / FramesPerSec
    FramesPerSecCounter = 0
    lFrameTimer = GetTickCount
End If
End Sub
 
Public Sub Connecting_Effect()
 
'Efectos
   
    'AlphaBlending en BOX'S
    If bConnectAb < 255 Then
    bConnectAb = bConnectAb + 5
    If bConnectAb > 245 Then bConnectAb = 255
    End If
   
    If aConnectAb < 30 Then
        aConnectAb = aConnectAb + 1
        If aConnectAb > 27 Then aConnectAb = 30
    End If
   
   
    Draw_FillBox 300, 281, 200, 130, D3DColorARGB(Val(aConnectAb), 128, 64, 0), D3DColorARGB(Val(aConnectAb), 255, 255, 255)
    Draw_FillBox 355, 392, 80, 25, D3DColorARGB(Val(aConnectAb), 128, 64, 0), D3DColorARGB(Val(aConnectAb), 255, 255, 255)
    
    Draw_FillBox 306, 310, 190, 10, D3DColorARGB(Val(aConnectAb), 128, 64, 0), D3DColorARGB(Val(aConnectAb), 255, 255, 255)
    Draw_FillBox 306, 342, 190, 10, D3DColorARGB(Val(aConnectAb), 128, 64, 0), D3DColorARGB(Val(aConnectAb), 255, 255, 255)
    
    
    
    Texto.Text_Draw 10, 10, "FPS:" & FramesPerSec, D3DColorXRGB(255, 255, 255)
    Texto.Text_Draw 306, 289, "Nombre:", D3DColorXRGB(255, 255, 255), , Val(bConnectAb)
    Texto.Text_Draw 306, 322, "Contraseña:", D3DColorXRGB(255, 255, 255), , Val(bConnectAb)
    Texto.Text_Draw 365, 392, "CONECTAR", D3DColorXRGB(255, 255, 255), , Val(bConnectAb)
    
    'Nick y password
    If FocoPasswd Then
        Texto.Text_Draw 306, 305, txtNombre, D3DColorXRGB(255, 255, 255), , Val(bConnectAb)
        Texto.Text_Draw 306, 337, txtPasswdAsteriscos, D3DColorXRGB(0, 255, 255), , Val(bConnectAb)
    Else
        Texto.Text_Draw 306, 305, txtNombre, D3DColorXRGB(0, 255, 255), , Val(bConnectAb)
        Texto.Text_Draw 306, 337, txtPasswdAsteriscos, D3DColorXRGB(255, 255, 255), , Val(bConnectAb)
    End If
    
'    Dim Lightt(3) As Long, LOGO As Grh
'    Lightt(0) = D3DColorXRGB(255, 255, 255)
'    Lightt(1) = D3DColorXRGB(255, 255, 255)
'    Lightt(2) = D3DColorXRGB(255, 255, 255)
'    Lightt(3) = D3DColorXRGB(255, 255, 255)
    
'    Call InitGrh(LOGO, 23662)
'    DDrawTransGrhtoSurface LOGO, 390, 200, 1, 1, Lightt(), , bConnectAb

End Sub

Public Sub Draw_FillBox(ByVal X As Integer, ByVal Y As Integer, ByVal Width As Integer, ByVal Height As Integer, Color As Long, outlinecolor As Long)
 
    Static box_rect As RECT
    Static Outline As RECT
    Static Rgb_List(3) As Long
    Static rgb_list2(3) As Long
    Static Vertex(3) As TLVERTEX
    Static Vertex2(3) As TLVERTEX
   
    Rgb_List(0) = Color
    Rgb_List(1) = Color
    Rgb_List(2) = Color
    Rgb_List(3) = Color
   
    rgb_list2(0) = outlinecolor
    rgb_list2(1) = outlinecolor
    rgb_list2(2) = outlinecolor
    rgb_list2(3) = outlinecolor
   
    With box_rect
        .Bottom = Y + Height - 1
        .Left = X + 1
        .Right = X + Width - 1
        .Top = Y + 1
    End With
   
    With Outline
        .Bottom = Y + Height
        .Left = X
        .Right = X + Width
        .Top = Y
    End With
   
   
    Geometry_Create_Box Vertex2(), Outline, Outline, rgb_list2(), 0, 0
    Geometry_Create_Box Vertex(), box_rect, box_rect, Rgb_List(), 0, 0
   
   
    d3ddevice.SetTexture 0, Nothing
    d3ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex2(0), Len(Vertex2(0))
    d3ddevice.DrawPrimitiveUP D3DPT_TRIANGLESTRIP, 2, Vertex(0), Len(Vertex(0))
 
   
End Sub

