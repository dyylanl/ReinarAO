VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTexManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private Const BYTES_PER_MB          As Long = 1048576
Private Const DEFAULT_MEMORY_TO_USE As Long = 64 ' In MB
Private Const HASH_TABLE_SIZE       As Long = 337
 
Private Type SURFACE_ENTRY_DYN
    filename As Integer
    UltimoAcceso As Long
    Texture As Direct3DTexture8
    Size As Long
    texture_width As Integer
    texture_height As Integer
End Type
 
Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type
 
Private TexList(HASH_TABLE_SIZE - 1) As HashNode
 
Private mD3D As D3DX8
Private device As Direct3DDevice8
 
' Parra was here (;
Private mGraphicsNumber As Long
Private mCurrentMemoryBytes As Long
Private mMaxMemoryBytes As Long
 
Private Sub Class_Terminate()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
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
End Sub
 
Public Property Get Surface(ByVal fileIndex As Long) As Direct3DTexture8
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Used to return a texture
'**************************************************************
End Property
 
Public Function GetTexture(ByVal filename As Integer, ByRef textwidth As Long, ByRef textheight As Long) As Direct3DTexture8
If filename = 0 Then
Debug.Print "0 GRH ATMPT TO BE LOADED"
Exit Function
End If
 
    Dim i As Long
    ' Search the index on the list
    With TexList(filename Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).filename = filename Then
                .SurfaceEntry(i).UltimoAcceso = GetTickCount
                textwidth = .SurfaceEntry(i).texture_width
                textheight = .SurfaceEntry(i).texture_height
                Set GetTexture = .SurfaceEntry(i).Texture
                Exit Function
            End If
        Next i
    End With
 
    'Not in memory, load it!
    Set GetTexture = CrearGrafico(filename, textwidth, textheight)
End Function
Private Function CrearGrafico(ByVal Archivo As Integer, ByRef texwidth As Long, ByRef textheight As Long) As Direct3DTexture8
On Error GoTo ErrHandler
    Dim surface_desc As D3DSURFACE_DESC
    Dim texture_info As D3DXIMAGE_INFO
    Dim index As Integer
    index = Archivo Mod HASH_TABLE_SIZE
    With TexList(index)
        .surfaceCount = .surfaceCount + 1
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        With .SurfaceEntry(.surfaceCount)
            'Nombre
            .filename = Archivo
           
            'Ultimo acceso
            .UltimoAcceso = GetTickCount
 
   Dim InfoHead As INFOHEADER
    Dim buffer() As Byte
   
    InfoHead = File_Find(App.Path & "\Graficos\Graficos.LAO", CStr(Archivo) & ".png")
   
    If InfoHead.lngFileSize <> 0 Then
                ' Parra was here (;
        mod_Compression.Extract_File_Memory Graphics, App.Path & "\Graficos\" & LCase$(CStr(Archivo) & ".png"), buffer()
           
        Set .Texture = D3DX.CreateTextureFromFileInMemoryEx(d3ddevice, buffer(0), UBound(buffer()) + 1, D3DX_DEFAULT, _
                                                            D3DX_DEFAULT, 0, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, _
                                                            D3DX_FILTER_POINT, D3DX_FILTER_NONE, _
                                                            &HFF000000, texture_info, ByVal 0)
        Erase buffer
    End If
               
            .Texture.GetLevelDesc 0, surface_desc
            .texture_width = texture_info.Width
            .texture_height = texture_info.Height
            .Size = surface_desc.Size
            texwidth = .texture_width
            textheight = .texture_height
            Set CrearGrafico = .Texture
            ' Parra was here (;
            'mFreeMemoryBytes = mFreeMemoryBytes + surface_desc.size
           
            ' Parra is back (;
            mGraphicsNumber = mGraphicsNumber + 1
            mCurrentMemoryBytes = mCurrentMemoryBytes + surface_desc.Size
        End With
    End With
   
    Do While mGraphicsNumber > 0 And mCurrentMemoryBytes > mMaxMemoryBytes
        If Not RemoveLRU() Then
            Exit Do
        End If
    Loop
Exit Function
ErrHandler:
Debug.Print "ERROR EN GRHLOAD>" & Archivo & ".png"
End Function
Private Function RemoveLRU() As Boolean
'**************************************************************
'Author: Juan Mart?n Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Removes the Least Recently Used surface to make some room for new ones
'**************************************************************
    Dim LRUi As Long
    Dim LRUj As Long
    Dim LRUtime As Long
    Dim i As Long
    Dim j As Long
    Dim surface_desc As D3DSURFACE_DESC
   
    LRUtime = GetTickCount
   
    'Check out through the whole list for the least recently used
    For i = 0 To HASH_TABLE_SIZE - 1
        With TexList(i)
            For j = 1 To .surfaceCount
                If LRUtime > .SurfaceEntry(j).UltimoAcceso Then
                    LRUi = i
                    LRUj = j
                    LRUtime = .SurfaceEntry(j).UltimoAcceso
                End If
            Next j
        End With
    Next i
   
    'Retrieve the surface desc
    Call TexList(LRUi).SurfaceEntry(LRUj).Texture.GetLevelDesc(0, surface_desc)
   
    ' Parra is back (;
    mCurrentMemoryBytes = mCurrentMemoryBytes - surface_desc.Size
   
    'Remove it
    Set TexList(LRUi).SurfaceEntry(LRUj).Texture = Nothing
    TexList(LRUi).SurfaceEntry(LRUj).filename = 0
   
    ' Parra is back (;
    mGraphicsNumber = mGraphicsNumber - 1
   
    'Move back the list (if necessary)
    If LRUj Then
        RemoveLRU = True
       
        With TexList(LRUi)
            For j = LRUj To .surfaceCount - 1
                .SurfaceEntry(j) = .SurfaceEntry(j + 1)
            Next j
           
            .surfaceCount = .surfaceCount - 1
            If .surfaceCount Then
                ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
            Else
                Erase .SurfaceEntry
            End If
        End With
    End If
   
    'Update the used bytes
    'mFreeMemoryBytes = mFreeMemoryBytes + surface_desc.size
End Function
