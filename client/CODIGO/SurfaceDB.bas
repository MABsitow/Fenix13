Attribute VB_Name = "SurfaceDB"
'**************************************************************
' clsSurfaceManDyn.cls - Inherits from clsSurfaceManager. Is designed to load
'surfaces dynamically without using more than an arbitrary amount of Mb.
'For removale it uses LRU, attempting to just keep in memory those surfaces
'that are actually usefull.
'
' Developed by Maraxus (Juan Martín Sotuyo Dodero - juansotuyo@hotmail.com)
' Last Modify Date: 3/06/2006
'**************************************************************

'**************************************************************
'This program is free software; you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation; either version 2 of the License, or
'any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program; if not, write to the Free Software
'Foundation, Inc., 59 Temple Place, Suite 330, Boston, MA  02111-1307  USA
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'**************************************************************

Option Explicit

Private Const BYTES_PER_MB As Long = 1048576                        '1Mb = 1024 Kb = 1024 * 1024 bytes = 1048576 bytes
Private Const MIN_MEMORY_TO_USE As Long = 4 * BYTES_PER_MB          '4 Mb
Private Const DEFAULT_MEMORY_TO_USE As Long = 16 * BYTES_PER_MB     '16 Mb

'Number of buckets in our hash table. Must be a nice prime number.
Const HASH_TABLE_SIZE As Long = 337

Private Type TYPE_SURFACE
        Width As Integer
        Height As Integer
        Surface As Long
        Size As Long
End Type

Private Type SURFACE_ENTRY_DYN
    fileIndex As Long
    lastAccess As Long
    Surface As TYPE_SURFACE
End Type

Private Type HashNode
    surfaceCount As Integer
    SurfaceEntry() As SURFACE_ENTRY_DYN
End Type

Private surfaceList(HASH_TABLE_SIZE - 1) As HashNode

Private maxBytesToUse As Long
Private usedBytes As Long

Private ResourcePath As String

Private Declare Function GetTickCount Lib "kernel32" () As Long

Public Sub Delete()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Clean up
'**************************************************************
    Dim i As Long
    Dim j As Long
    
    'Destroy every surface in memory
    For i = 0 To HASH_TABLE_SIZE - 1
        With surfaceList(i)
            For j = 1 To .surfaceCount
                Dim tex As TYPE_SURFACE
                
                .SurfaceEntry(j).Surface = tex
            Next j
            
            'Destroy the arrays
            Erase .SurfaceEntry
        End With
    Next i
End Sub

Public Sub Initialize(ByVal graphicPath As String, Optional ByVal maxMemoryUsageInMb As Long = -1)
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 29/07/2012 - ^[GS]^
'Initializes the manager
'**************************************************************
    
    usedBytes = 0
    maxBytesToUse = MIN_MEMORY_TO_USE
    
    ResourcePath = graphicPath
    
    If maxMemoryUsageInMb = -1 Then
        maxBytesToUse = DEFAULT_MEMORY_TO_USE   ' 16 Mb by default
    ElseIf maxMemoryUsageInMb * BYTES_PER_MB < MIN_MEMORY_TO_USE Then
        maxBytesToUse = MIN_MEMORY_TO_USE       ' 4 Mb is the minimum allowed
    Else
        maxBytesToUse = maxMemoryUsageInMb * BYTES_PER_MB
    End If
End Sub

Public Property Get Surface(ByVal fileIndex As Long) As Texture
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Retrieves the requested texture
'**************************************************************
    Dim i As Long
    
    ' Search the index on the list
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        For i = 1 To .surfaceCount
            If .SurfaceEntry(i).fileIndex = fileIndex Then
                .SurfaceEntry(i).lastAccess = GetTickCount
                
                Dim tex As Texture
                tex.Ptr = .SurfaceEntry(i).Surface.Surface
                tex.Width = .SurfaceEntry(i).Surface.Width
                tex.Height = .SurfaceEntry(i).Surface.Height
                
                Surface = tex
                Exit Property
            End If
        Next i
    End With
    
    'Not in memory, load it!
    Surface = LoadSurface(fileIndex)
End Property

Private Function LoadSurface(ByVal fileIndex As Long) As Texture
'**************************************************************
'Author: Nicolas Matias Gonzalez (NIGO)
'Last Modify Date: 09/10/2012 - ^[GS]^
'Loads the surface named fileIndex + ".bmp" and inserts it to the
'surface list in the listIndex position
'**************************************************************
On Error GoTo ErrHandler
    
    Dim tex As Texture
    Dim newSurface As SURFACE_ENTRY_DYN
    Dim Image As TYPE_VIDEO_IMAGE
        
    With newSurface
        .fileIndex = fileIndex
        
        'Set last access time (if we didn't we would reckon this texture as the one lru)
        .lastAccess = GetTickCount

        Image = Video.CreateImageFromFilename(DirGraficos & fileIndex & ".png", False, , TEXTURE_FORMAT_RGBA8)

        '.Surface.Surface = Video.CreateTexture2DFromMemory(Image.vX, Image.vY, False, 1, TEXTURE_FORMAT_RGBA8, TEXTURE_FLAG_NONE, BGFX.Copy(Image.vData, Image.vX * Image.vY * Image.vComponent))
        
        .Surface.Surface = Image.mHandle
        
        .Surface.Width = Image.mX
        .Surface.Height = Image.mY
        
    End With

    'Insert surface to the list
    With surfaceList(fileIndex Mod HASH_TABLE_SIZE)
        .surfaceCount = .surfaceCount + 1
        
        ReDim Preserve .SurfaceEntry(1 To .surfaceCount) As SURFACE_ENTRY_DYN
        
        .SurfaceEntry(.surfaceCount) = newSurface
        
        tex.Ptr = newSurface.Surface.Surface
        tex.Width = newSurface.Surface.Width
        tex.Height = newSurface.Surface.Height
        
        LoadSurface = tex
    End With
    
    'Update used bytes
    usedBytes = usedBytes + newSurface.Surface.Size
    
    'Check if we have exceeded our allowed share of memory usage
    Do While usedBytes > maxBytesToUse
        'Remove a file. If no file could be removed we continue, if the file was previous to our surface we update the index
        If Not RemoveLRU() Then
            Exit Do
        End If
    Loop
Exit Function

ErrHandler:

End Function

Private Function RemoveLRU() As Boolean
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/06/2006
'Removes the Least Recently Used surface to make some room for new ones
'**************************************************************
    Dim LRUi As Long
    Dim LRUj As Long
    Dim LRUtime As Long
    Dim i As Long
    Dim j As Long
    Dim Size As Long
    
    LRUtime = GetTickCount
    
    'TODO: REWRITE
    
    'Check out through the whole list for the least recently used
    For i = 0 To HASH_TABLE_SIZE - 1
        With surfaceList(i)
            For j = 1 To .surfaceCount
                If LRUtime > .SurfaceEntry(j).lastAccess Then
                    LRUi = i
                    LRUj = j
                    LRUtime = .SurfaceEntry(j).lastAccess
                End If
            Next j
        End With
    Next i
    
    If LRUj Then
        RemoveLRU = True
        'Remove it
        'surfaceList(LRUi).SurfaceEntry(LRUj).Surface.Surface = -1
        surfaceList(LRUi).SurfaceEntry(LRUj).fileIndex = 0
        
        'Update the used bytes
        usedBytes = usedBytes - surfaceList(LRUi).SurfaceEntry(LRUj).Surface.Size
        
        'Move back the list (if necessary)
        With surfaceList(LRUi)
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
End Function
