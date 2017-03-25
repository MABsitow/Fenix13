VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSurfaceDB"
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
    size As Long
    texture_width As Integer
    texture_height As Integer
End Type
 
Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type
 
Private TexList(HASH_TABLE_SIZE - 1) As HashNode
 
Private mD3D As D3DX8
Private mDevice As Direct3DDevice8
 
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
 
'CSEH: ErrLog
Public Function GetTexture(ByVal filename As Integer, ByRef textwidth As Long, ByRef textheight As Long) As Direct3DTexture8
    '<EhHeader>
    On Error GoTo GetTexture_Err
    '</EhHeader>
100 If filename = 0 Then
105 Debug.Print "0 GRH ATMPT TO BE LOADED"
    Exit Function
    End If
 
        Dim i As Long
        ' Search the index on the list
110     With TexList(filename Mod HASH_TABLE_SIZE)
115         For i = 1 To .surfaceCount
120             If .SurfaceEntry(i).filename = filename Then
125                 .SurfaceEntry(i).UltimoAcceso = GetTickCount
130                 textwidth = .SurfaceEntry(i).texture_width
135                 textheight = .SurfaceEntry(i).texture_height
140                 Set GetTexture = .SurfaceEntry(i).Texture
                    Exit Function
                End If
145         Next i
        End With
 
        'Not in memory, load it!
150     Set GetTexture = CrearGrafico(filename, textwidth, textheight)
    '<EhFooter>
    Exit Function

GetTexture_Err:
        Call LogError("Error en GetTexture: " & Erl & " - " & Err.Description)
    '</EhFooter>
End Function
 
Public Function Init(ByRef D3D8 As D3DX8, ByRef Device As Direct3DDevice8, Optional ByVal MaxMemory As Long = DEFAULT_MEMORY_TO_USE) As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 5/04/2005
'
'**************************************************************
   
    'Seteamos el objeto
    Set mD3D = D3D8
    Set mDevice = Device
   
    ' Parra is back (;
    mGraphicsNumber = 0
    mCurrentMemoryBytes = 0
    mMaxMemoryBytes = MaxMemory * BYTES_PER_MB
   
    Init = True
End Function

'CSEH: ErrLog
Private Function CrearGrafico(ByVal Archivo As Integer, ByRef TexWidth As Long, ByRef textheight As Long) As Direct3DTexture8
    '<EhHeader>
    On Error GoTo CrearGrafico_Err
    '</EhHeader>
        Dim surface_desc As D3DSURFACE_DESC
        Dim texture_info As D3DXIMAGE_INFO
        Dim Index As Integer
100     Index = Archivo Mod HASH_TABLE_SIZE
105     With TexList(Index)
110         .surfaceCount = .surfaceCount + 1
115         ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
120         With .SurfaceEntry(.surfaceCount)
                'Nombre
125             .filename = Archivo
           
                'Ultimo acceso
130             .UltimoAcceso = GetTickCount
 
   
                ' Parra was here (;
135             Set .Texture = mD3D.CreateTextureFromFileEx(mDevice, DirGraficos & Archivo & ".png", D3DX_DEFAULT, _
                                                                D3DX_DEFAULT, 0, 0, D3DFMT_A8R8G8B8, D3DPOOL_MANAGED, _
                                                                D3DX_FILTER_POINT, D3DX_FILTER_NONE, _
                                                                &HFF000000, texture_info, ByVal 0)
140             .Texture.GetLevelDesc 0, surface_desc
145             .texture_width = texture_info.Width
150             .texture_height = texture_info.Height
155             .size = surface_desc.size
160             TexWidth = .texture_width
165             textheight = .texture_height
170             Set CrearGrafico = .Texture
                ' Parra was here (;
                'mFreeMemoryBytes = mFreeMemoryBytes + surface_desc.size
           
                ' Parra is back (;
175             mGraphicsNumber = mGraphicsNumber + 1
180             mCurrentMemoryBytes = mCurrentMemoryBytes + surface_desc.size
            End With
        End With
   
185     Do While mGraphicsNumber > 0 And mCurrentMemoryBytes > mMaxMemoryBytes
190         If Not RemoveLRU() Then
                Exit Do
            End If
        Loop
    
    '<EhFooter>
    Exit Function

CrearGrafico_Err:
        Call LogError("Error en CrearGrafico: " & Erl & " - " & Err.Description)
    '</EhFooter>
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
    mCurrentMemoryBytes = mCurrentMemoryBytes - surface_desc.size
   
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

