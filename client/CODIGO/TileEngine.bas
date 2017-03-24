Attribute VB_Name = "Mod_TileEngine"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Public SurfaceDB As clsSurfaceDB
Private SpriteBatch As clsBatch

Public DirectX As DirectX8
Public DirectD3D8 As D3DX8
Public DirectD3D As Direct3D8
Public DirectDevice As Direct3DDevice8

Public Type TLVERTEX
    X As Single
    Y As Single
    Z As Single
    
    Color As Long

    Tu As Single
    Tv As Single
End Type

Private MainUITex As Direct3DTexture8
Private LoginTex  As Direct3DTexture8

Private Viewport As D3DVIEWPORT8
Private Projection As D3DMATRIX
Private View As D3DMATRIX
'Private Camera As D3DVECTOR

'Map sizes in tiles
Public Const XMaxMapSize As Byte = 100
Public Const XMinMapSize As Byte = 1
Public Const YMaxMapSize As Byte = 100
Public Const YMinMapSize As Byte = 1

Private Const GrhFogata As Integer = 1521

''
'Sets a Grh animation to loop indefinitely.
Private Const INFINITE_LOOPS As Integer = -1


'Encabezado bmp
Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

'Info del encabezado del bmp
Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

'Posicion en un mapa
Public Type Position
    X As Long
    Y As Long
End Type

'Posicion en el Mundo
Public Type WorldPos
    Map As Integer
    X As Integer
    Y As Integer
End Type

'Contiene info acerca de donde se puede encontrar un grh tamaño y animacion
Public Type GrhData
    sX As Integer
    sY As Integer
    
    FileNum As Long
    
    pixelWidth As Integer
    pixelHeight As Integer
    
    TileWidth As Single
    TileHeight As Single
    
    NumFrames As Integer
    Frames(1 To 25) As Integer
    
    Speed As Single
End Type

'apunta a una estructura grhdata y mantiene la animacion
Public Type Grh
    GrhIndex As Integer
    FrameCounter As Single
    Speed As Single
    Started As Byte
    Loops As Integer
End Type

'Lista de cuerpos
Public Type BodyData
    Walk(E_Heading.NORTH To E_Heading.WEST) As Grh
    HeadOffset As Position
End Type

'Lista de cabezas
Public Type HeadData
    Head(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de las armas
Type WeaponAnimData
    WeaponWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type

'Lista de las animaciones de los escudos
Type ShieldAnimData
    ShieldWalk(E_Heading.NORTH To E_Heading.WEST) As Grh
End Type


'Apariencia del personaje
Public Type Char
    Active As Byte
    Heading As E_Heading
    Pos As Position
    
    iHead As Integer
    iBody As Integer
    Body As BodyData
    Head As HeadData
    Casco As HeadData
    Arma As WeaponAnimData
    Escudo As ShieldAnimData
    UsandoArma As Boolean
    
    fX As Grh
    FxIndex As Integer
    
    Criminal As Byte
    
    Nombre As String
    NombreOffset As Integer
    
    scrollDirectionX As Integer
    scrollDirectionY As Integer
    
    Moving As Byte
    MoveOffsetX As Single
    MoveOffsetY As Single
    
    pie As Boolean
    muerto As Boolean
    invisible As Boolean
    priv As Byte
End Type

'Info de un objeto
Public Type Obj
    OBJIndex As Integer
    amount As Integer
End Type

'Tipo de las celdas del mapa
Public Type MapBlock
    Graphic(1 To 4) As Grh
    CharIndex As Integer
    ObjGrh As Grh
    
    NPCIndex As Integer
    OBJInfo As Obj
    TileExit As WorldPos
    Blocked As Byte
    
    Trigger As Integer
End Type

'Info de cada mapa
Public Type MapInfo
    Music As String
    Name As String
    StartPos As WorldPos
    MapVersion As Integer
End Type

Public Const DegreeToRadian As Single = 0.01745329251994 'Pi / 180


Public AmbientColor(3) As Long

Public IniPath As String

'Bordes del mapa
Public MinXBorder As Byte
Public MaxXBorder As Byte
Public MinYBorder As Byte
Public MaxYBorder As Byte

'Status del user
Public CurMap As Integer 'Mapa actual
Public UserIndex As Integer
Public UserMoving As Byte
Public UserBody As Integer
Public UserHead As Integer
Public UserPos As Position 'Posicion
Public AddtoUserPos As Position 'Si se mueve
Public UserCharIndex As Integer

Public EngineRun As Boolean

Public FPS As Long
Public FramesPerSecCounter As Long
Private fpsLastCheck As Long

'Tamaño del la vista en Tiles
Private WindowTileWidth As Integer
Private WindowTileHeight As Integer

Private HalfWindowTileWidth As Integer
Private HalfWindowTileHeight As Integer

'Offset del desde 0,0 del main view
Private MainViewTop As Integer
Private MainViewLeft As Integer

'Cuantos tiles el engine mete en el BUFFER cuando
'dibuja el mapa. Ojo un tamaño muy grande puede
'volver el engine muy lento
Public TileBufferSize As Integer

'Private TileBufferPixelOffsetX As Integer
'Private TileBufferPixelOffsetY As Integer

'Tamaño de los tiles en pixels
Public TilePixelHeight As Integer
Public TilePixelWidth As Integer

'Number of pixels the engine scrolls per frame. MUST divide evenly into pixels per tile
Public ScrollPixelsPerFrameX As Integer
Public ScrollPixelsPerFrameY As Integer

Dim timerElapsedTime As Single
Dim timerTicksPerFrame As Single
Dim engineBaseSpeed As Single

Public NumChars As Integer
Public LastChar As Integer
Public NumWeaponAnims As Integer

Private MainView        As Long 'ID
Private MainViewWidth   As Integer
Private MainViewHeight  As Integer

Private MouseTileX As Byte
Private MouseTileY As Byte

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Graficos¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public GrhData() As GrhData 'Guarda todos los grh
Public BodyData() As BodyData
Public HeadData() As HeadData
Public FxData() As tIndiceFx
Public WeaponAnimData() As WeaponAnimData
Public ShieldAnimData() As ShieldAnimData
Public CascoAnimData() As HeadData
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿Mapa?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?
Public MapData() As MapBlock ' Mapa
Public MapInfo As MapInfo ' Info acerca del mapa en uso
'¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?¿?

Public bRain        As Boolean 'está raineando?
Public bTecho       As Boolean 'hay techo?

Public charlist(1 To 10000) As Char

' Used by GetTextExtentPoint32
Private Type size
    cx As Long
    cy As Long
End Type

Public Enum PlayLoop
    plNone = 0
    plLluviain = 1
    plLluviaout = 2
End Enum

'Very percise counter 64bit system counter
Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long

Sub CargarCabezas()
    Dim N As Integer
    Dim i As Long
    Dim Numheads As Integer
    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cabezas.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Numheads
    
    'Resize array
    ReDim HeadData(0 To Numheads) As HeadData
    ReDim Miscabezas(0 To Numheads) As tIndiceCabeza
    
    For i = 1 To Numheads
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(HeadData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(HeadData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(HeadData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(HeadData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCascos()
    Dim N As Integer
    Dim i As Long
    Dim NumCascos As Integer

    Dim Miscabezas() As tIndiceCabeza
    
    N = FreeFile()
    Open App.path & "\init\Cascos.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCascos
    
    'Resize array
    ReDim CascoAnimData(0 To NumCascos) As HeadData
    ReDim Miscabezas(0 To NumCascos) As tIndiceCabeza
    
    For i = 1 To NumCascos
        Get #N, , Miscabezas(i)
        
        If Miscabezas(i).Head(1) Then
            Call InitGrh(CascoAnimData(i).Head(1), Miscabezas(i).Head(1), 0)
            Call InitGrh(CascoAnimData(i).Head(2), Miscabezas(i).Head(2), 0)
            Call InitGrh(CascoAnimData(i).Head(3), Miscabezas(i).Head(3), 0)
            Call InitGrh(CascoAnimData(i).Head(4), Miscabezas(i).Head(4), 0)
        End If
    Next i
    
    Close #N
End Sub

Sub CargarCuerpos()
    Dim N As Integer
    Dim i As Long
    Dim NumCuerpos As Integer
    Dim MisCuerpos() As tIndiceCuerpo
    
    N = FreeFile()
    Open App.path & "\init\Personajes.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumCuerpos
    
    'Resize array
    ReDim BodyData(0 To NumCuerpos) As BodyData
    ReDim MisCuerpos(0 To NumCuerpos) As tIndiceCuerpo
    
    For i = 1 To NumCuerpos
        Get #N, , MisCuerpos(i)
        
        If MisCuerpos(i).Body(1) Then
            InitGrh BodyData(i).Walk(1), MisCuerpos(i).Body(1), 0
            InitGrh BodyData(i).Walk(2), MisCuerpos(i).Body(2), 0
            InitGrh BodyData(i).Walk(3), MisCuerpos(i).Body(3), 0
            InitGrh BodyData(i).Walk(4), MisCuerpos(i).Body(4), 0
            
            BodyData(i).HeadOffset.X = MisCuerpos(i).HeadOffsetX
            BodyData(i).HeadOffset.Y = MisCuerpos(i).HeadOffsetY
        End If
    Next i
    
    Close #N
End Sub

Sub CargarFxs()
    Dim N As Integer
    Dim i As Long
    Dim NumFxs As Integer
    
    N = FreeFile()
    Open App.path & "\init\Fxs.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumFxs
    
    'Resize array
    ReDim FxData(1 To NumFxs) As tIndiceFx
    
    For i = 1 To NumFxs
        Get #N, , FxData(i)
    Next i
    
    Close #N
End Sub

Sub CargarTips()
    Dim N As Integer
    Dim i As Long
    Dim NumTips As Integer
    
    N = FreeFile
    Open App.path & "\init\Tips.ayu" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , NumTips
    
    'Resize array
    ReDim Tips(1 To NumTips) As String * 255
    
    For i = 1 To NumTips
        Get #N, , Tips(i)
    Next i
    
    Close #N
End Sub

Sub CargarArrayLluvia()
    Dim N As Integer
    Dim i As Long
    Dim Nu As Integer
    
    N = FreeFile()
    Open App.path & "\init\fk.ind" For Binary Access Read As #N
    
    'cabecera
    Get #N, , MiCabecera
    
    'num de cabezas
    Get #N, , Nu
    
    'Resize array
    ReDim bLluvia(1 To Nu) As Byte
    
    For i = 1 To Nu
        Get #N, , bLluvia(i)
    Next i
    
    Close #N
End Sub

Sub ConvertCPtoTP(ByVal viewPortX As Integer, ByVal viewPortY As Integer, ByRef TX As Byte, ByRef TY As Byte)
'******************************************
'Converts where the mouse is in the main window to a tile position. MUST be called eveytime the mouse moves.
'******************************************
    TX = UserPos.X + viewPortX \ TilePixelWidth - WindowTileWidth \ 2
    TY = UserPos.Y + viewPortY \ TilePixelHeight - WindowTileHeight \ 2
End Sub

Sub MakeChar(ByVal CharIndex As Integer, ByVal Body As Integer, ByVal Head As Integer, ByVal Heading As Byte, ByVal X As Integer, ByVal Y As Integer, ByVal Arma As Integer, ByVal Escudo As Integer, ByVal Casco As Integer)
On Error Resume Next
    'Apuntamos al ultimo Char
    If CharIndex > LastChar Then LastChar = CharIndex
    
    With charlist(CharIndex)
        'If the char wasn't allready active (we are rewritting it) don't increase char count
        If .Active = 0 Then _
            NumChars = NumChars + 1
        
        If Arma = 0 Then Arma = 2
        If Escudo = 0 Then Escudo = 2
        If Casco = 0 Then Casco = 2
        
        .iHead = Head
        .iBody = Body
        .Head = HeadData(Head)
        .Body = BodyData(Body)
        .Arma = WeaponAnimData(Arma)
        
        .Escudo = ShieldAnimData(Escudo)
        .Casco = CascoAnimData(Casco)
        
        .Heading = Heading
        
        'Reset moving stats
        .Moving = 0
        .MoveOffsetX = 0
        .MoveOffsetY = 0
        
        'Update position
        .Pos.X = X
        .Pos.Y = Y
        
        'Make active
        .Active = 1
    End With
    
    'Plot on map
    MapData(X, Y).CharIndex = CharIndex
End Sub

Sub ResetCharInfo(ByVal CharIndex As Integer)
    With charlist(CharIndex)
        .Active = 0
        .Criminal = 0
        .FxIndex = 0
        .invisible = False
        
        .Moving = 0
        .muerto = False
        .Nombre = ""
        .NombreOffset = 0
        .pie = False
        .Pos.X = 0
        .Pos.Y = 0
        .UsandoArma = False
    End With
End Sub

Sub EraseChar(ByVal CharIndex As Integer)
'*****************************************************************
'Erases a character from CharList and map
'*****************************************************************
On Error Resume Next
    charlist(CharIndex).Active = 0
    
    'Update lastchar
    If CharIndex = LastChar Then
        Do Until charlist(LastChar).Active = 1
            LastChar = LastChar - 1
            If LastChar = 0 Then Exit Do
        Loop
    End If
    
    MapData(charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y).CharIndex = 0
    
    'Remove char's dialog
    Call Dialogos.RemoveDialog(CharIndex)
    
    Call ResetCharInfo(CharIndex)
    
    'Update NumChars
    NumChars = NumChars - 1
End Sub

'CSEH: ErrLog
Public Sub InitGrh(ByRef Grh As Grh, ByVal GrhIndex As Integer, Optional ByVal Started As Byte = 2)
    '*****************************************************************
    'Sets up a grh. MUST be done before rendering
    '*****************************************************************
    '<EhHeader>
    On Error GoTo InitGrh_Err
    '</EhHeader>
        If GrhIndex = 0 Then Exit Sub

100     Grh.GrhIndex = GrhIndex
    
105     If Started = 2 Then
110         If GrhData(Grh.GrhIndex).NumFrames > 1 Then
115             Grh.Started = 1
            Else
120             Grh.Started = 0
            End If
        Else
            'Make sure the graphic can be started
125         If GrhData(Grh.GrhIndex).NumFrames = 1 Then Started = 0
130         Grh.Started = Started
        End If
    
    
135     If Grh.Started Then
140         Grh.Loops = INFINITE_LOOPS
        Else
145         Grh.Loops = 0
        End If
    
150     Grh.FrameCounter = 1
155     Grh.Speed = GrhData(Grh.GrhIndex).Speed
    '<EhFooter>
    Exit Sub

InitGrh_Err:
        Call LogError("Error en InitGrh: " & Erl & " - " & Err.Description)
    '</EhFooter>
End Sub

Sub MoveCharbyHead(ByVal CharIndex As Integer, ByVal nHeading As E_Heading)
'*****************************************************************
'Starts the movement of a character in nHeading direction
'*****************************************************************
    Dim addX As Integer
    Dim addY As Integer
    Dim X As Integer
    Dim Y As Integer
    Dim nX As Integer
    Dim nY As Integer
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        'Figure out which way to move
        Select Case nHeading
            Case E_Heading.NORTH
                addY = -1
        
            Case E_Heading.EAST
                addX = 1
        
            Case E_Heading.SOUTH
                addY = 1
            
            Case E_Heading.WEST
                addX = -1
        End Select
        
        nX = X + addX
        nY = Y + addY
        
        MapData(nX, nY).CharIndex = CharIndex
        .Pos.X = nX
        .Pos.Y = nY
        MapData(X, Y).CharIndex = 0
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = addX
        .scrollDirectionY = addY
    End With
    
    If UserEstado = 0 Then Call DoPasosFx(CharIndex)
    
    'areas viejos
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        If CharIndex <> UserCharIndex Then
            Call EraseChar(CharIndex)
        End If
    End If
End Sub

Public Sub DoFogataFx()
    Dim location As Position
    
    If bFogata Then
        bFogata = HayFogata(location)
        If Not bFogata Then
            'call 'audio.StopWave(FogataBufferIndex)
            FogataBufferIndex = 0
        End If
    Else
        bFogata = HayFogata(location)
        'If bFogata And FogataBufferIndex = 0 Then FogataBufferIndex = 'audio.PlayWave("fuego.wav", location.X, location.Y, LoopStyle.Enabled)
    End If
End Sub

Private Function EstaPCarea(ByVal CharIndex As Integer) As Boolean
    With charlist(CharIndex).Pos
        EstaPCarea = .X > UserPos.X - MinXBorder And .X < UserPos.X + MinXBorder And .Y > UserPos.Y - MinYBorder And .Y < UserPos.Y + MinYBorder
    End With
End Function

Sub DoPasosFx(ByVal CharIndex As Integer)
    If Not UserNavegando Then
        With charlist(CharIndex)
            If Not .muerto And EstaPCarea(CharIndex) And (.priv = 0 Or .priv > 5) Then
                .pie = Not .pie
                
                If .pie Then
                    'call 'audio.PlayWave(SND_PASOS1, .Pos.X, .Pos.Y)
                Else
                    'call 'audio.PlayWave(SND_PASOS2, .Pos.X, .Pos.Y)
                End If
            End If
        End With
    Else
' TODO : Actually we would have to check if the CharIndex char is in the water or not....
        'call 'audio.PlayWave(SND_NAVEGANDO, charlist(CharIndex).Pos.X, charlist(CharIndex).Pos.Y)
    End If
End Sub

Sub MoveCharbyPos(ByVal CharIndex As Integer, ByVal nX As Integer, ByVal nY As Integer)
    Dim X As Integer
    Dim Y As Integer
    Dim addX As Integer
    Dim addY As Integer
    Dim nHeading As E_Heading
    
    With charlist(CharIndex)
        X = .Pos.X
        Y = .Pos.Y
        
        MapData(X, Y).CharIndex = 0
        
        addX = nX - X
        addY = nY - Y
        
        If Sgn(addX) = 1 Then
            nHeading = E_Heading.EAST
        ElseIf Sgn(addX) = -1 Then
            nHeading = E_Heading.WEST
        ElseIf Sgn(addY) = -1 Then
            nHeading = E_Heading.NORTH
        ElseIf Sgn(addY) = 1 Then
            nHeading = E_Heading.SOUTH
        End If
        
        MapData(nX, nY).CharIndex = CharIndex
        
        .Pos.X = nX
        .Pos.Y = nY
        
        .MoveOffsetX = -1 * (TilePixelWidth * addX)
        .MoveOffsetY = -1 * (TilePixelHeight * addY)
        
        .Moving = 1
        .Heading = nHeading
        
        .scrollDirectionX = Sgn(addX)
        .scrollDirectionY = Sgn(addY)
        
        'parche para que no medite cuando camina
        If .FxIndex = FxMeditar.CHICO Or .FxIndex = FxMeditar.GRANDE Or .FxIndex = FxMeditar.MEDIANO Or .FxIndex = FxMeditar.XGRANDE Or .FxIndex = FxMeditar.XXGRANDE Then
            .FxIndex = 0
        End If
    End With
    
    If Not EstaPCarea(CharIndex) Then Call Dialogos.RemoveDialog(CharIndex)
    
    If (nY < MinLimiteY) Or (nY > MaxLimiteY) Or (nX < MinLimiteX) Or (nX > MaxLimiteX) Then
        Call EraseChar(CharIndex)
    End If
End Sub

Sub MoveCamera(ByVal nHeading As E_Heading)

    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            'Camera.Y = Camera.Y - 1
        Case E_Heading.EAST
            'Camera.X = Camera.X + 1
        Case E_Heading.SOUTH
            'Camera.Y = Camera.Y + 1
        Case E_Heading.WEST
            'Camera.X = Camera.X - 1
    End Select
    
    
    'Check to see if its out of bounds
    'If Camera.X < MinXBorder Or Camera.X > MaxXBorder Or Camera.Y < MinYBorder Or Camera.Y > MaxYBorder Then
    '    Exit Sub
    'Else
    
    '    With At
    '        .X = Camera.X
    '        .Y = Camera.Y
    '        .Z = 0
    '    End With
        
    '    Call D3DXMATH_MATRIX.D3DXMatrixLookAtLH(View, Camera, At, Up)
    '    Call DirectDevice.SetTransform(D3DTS_VIEW, View)
        
   ' End If
End Sub

Sub MoveScreen(ByVal nHeading As E_Heading)
'******************************************
'Starts the screen moving in a direction
'******************************************
    Dim X As Integer
    Dim Y As Integer
    Dim TX As Integer
    Dim TY As Integer
    
    'Figure out which way to move
    Select Case nHeading
        Case E_Heading.NORTH
            Y = -1
        Case E_Heading.EAST
            X = 1
        Case E_Heading.SOUTH
            Y = 1
        Case E_Heading.WEST
            X = -1
    End Select
    
    'Fill temp pos
    TX = UserPos.X + X
    TY = UserPos.Y + Y
    
    'Check to see if its out of bounds
    If TX < MinXBorder Or TX > MaxXBorder Or TY < MinYBorder Or TY > MaxYBorder Then
        Exit Sub
    Else
        'Start moving... MainLoop does the rest
        AddtoUserPos.X = X
        UserPos.X = TX
        AddtoUserPos.Y = Y
        UserPos.Y = TY
        UserMoving = 1
        
        bTecho = IIf(MapData(UserPos.X, UserPos.Y).Trigger = 1 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 2 Or _
                MapData(UserPos.X, UserPos.Y).Trigger = 4, True, False)
        
    End If
End Sub

Private Function HayFogata(ByRef location As Position) As Boolean
    Dim j As Long
    Dim k As Long
    
    For j = UserPos.X - 8 To UserPos.X + 8
        For k = UserPos.Y - 6 To UserPos.Y + 6
            If InMapBounds(j, k) Then
                If MapData(j, k).ObjGrh.GrhIndex = GrhFogata Then
                    location.X = j
                    location.Y = k
                    
                    HayFogata = True
                    Exit Function
                End If
            End If
        Next k
    Next j
End Function

Function NextOpenChar() As Integer
'*****************************************************************
'Finds next open char slot in CharList
'*****************************************************************
    Dim loopC As Long
    Dim Dale As Boolean
    
    loopC = 1
    Do While charlist(loopC).Active And Dale
        loopC = loopC + 1
        Dale = (loopC <= UBound(charlist))
    Loop
    
    NextOpenChar = loopC
End Function

''
' Loads grh data using the new file format.
'
' @return   True if the load was successfull, False otherwise.

Private Function LoadGrhData() As Boolean
On Error GoTo errorHandler
    Dim Grh As Long
    Dim Frame As Long
    Dim grhCount As Long
    Dim handle As Integer
    Dim fileVersion As Long
    
    'Open files
    handle = FreeFile()
    
    Open IniPath & "Graficos.ind" For Binary Access Read As handle
    Seek #1, 1
    
    'Get file version
    Get handle, , fileVersion
    
    'Get number of grhs
    Get handle, , grhCount
    
    'Resize arrays
    ReDim GrhData(1 To grhCount) As GrhData
    
    Get handle, , Grh
    
    Do Until Grh <= 0
        
        With GrhData(Grh)
            
            'Get number of frames
            Get handle, , .NumFrames
            If .NumFrames <= 0 Then GoTo errorHandler
            
            'ReDim .Frames(1 To GrhData(Grh).NumFrames)
            
            If .NumFrames > 1 Then
                'Read a animation GRH set
                For Frame = 1 To .NumFrames
                    Get handle, , .Frames(Frame)
                    If .Frames(Frame) <= 0 Or .Frames(Frame) > grhCount Then
                        GoTo errorHandler
                    End If
                Next Frame
                
                Get handle, , .Speed
                
                If .Speed <= 0 Then GoTo errorHandler
                
                'Compute width and height
                .pixelHeight = GrhData(.Frames(1)).pixelHeight
                If .pixelHeight <= 0 Then GoTo errorHandler
                
                .pixelWidth = GrhData(.Frames(1)).pixelWidth
                If .pixelWidth <= 0 Then GoTo errorHandler
                
                .TileWidth = GrhData(.Frames(1)).TileWidth
                If .TileWidth <= 0 Then GoTo errorHandler
                
                .TileHeight = GrhData(.Frames(1)).TileHeight
                If .TileHeight <= 0 Then GoTo errorHandler
            Else
                'Read in normal GRH data
                Get handle, , .FileNum
                If .FileNum <= 0 Then GoTo errorHandler
                
                Get handle, , GrhData(Grh).sX
                If .sX < 0 Then GoTo errorHandler
                
                Get handle, , .sY
                If .sY < 0 Then GoTo errorHandler
                
                Get handle, , .pixelWidth
                If .pixelWidth <= 0 Then GoTo errorHandler
                
                Get handle, , .pixelHeight
                If .pixelHeight <= 0 Then GoTo errorHandler
                
                'Compute width and height
                .TileWidth = .pixelWidth / TilePixelHeight
                .TileHeight = .pixelHeight / TilePixelWidth
                
                .Frames(1) = Grh
            End If
        End With
        
        Get handle, , Grh
    Loop
    
    Close handle
    
    LoadGrhData = True
Exit Function

errorHandler:
    LoadGrhData = False
End Function

Function LegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is legal
'*****************************************************************
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    '¿Hay un personaje?
    If MapData(X, Y).CharIndex > 0 Then
        Exit Function
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    LegalPos = True
End Function

Function MoveToLegalPos(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Author: ZaMa
'Last Modify Date: 01/08/2009
'Checks to see if a tile position is legal, including if there is a casper in the tile
'10/05/2009: ZaMa - Now you can't change position with a casper which is in the shore.
'01/08/2009: ZaMa - Now invisible admins can't change position with caspers.
'*****************************************************************
    Dim CharIndex As Integer
    
    'Limites del mapa
    If X < MinXBorder Or X > MaxXBorder Or Y < MinYBorder Or Y > MaxYBorder Then
        Exit Function
    End If
    
    'Tile Bloqueado?
    If MapData(X, Y).Blocked = 1 Then
        Exit Function
    End If
    
    CharIndex = MapData(X, Y).CharIndex
    '¿Hay un personaje?
    If CharIndex > 0 Then
    
        If MapData(UserPos.X, UserPos.Y).Blocked = 1 Then
            Exit Function
        End If
        
        With charlist(CharIndex)
            ' Si no es casper, no puede pasar
            If .iHead <> CASPER_HEAD And .iBody <> FRAGATA_FANTASMAL Then
                Exit Function
            Else
                ' No puedo intercambiar con un casper que este en la orilla (Lado tierra)
                If HayAgua(UserPos.X, UserPos.Y) Then
                    If Not HayAgua(X, Y) Then Exit Function
                Else
                    ' No puedo intercambiar con un casper que este en la orilla (Lado agua)
                    If HayAgua(X, Y) Then Exit Function
                End If
                
                ' Los admins no pueden intercambiar pos con caspers cuando estan invisibles
                If charlist(UserCharIndex).priv > 0 And charlist(UserCharIndex).priv < 6 Then
                    If charlist(UserCharIndex).invisible = True Then Exit Function
                End If
            End If
        End With
    End If
   
    If UserNavegando <> HayAgua(X, Y) Then
        Exit Function
    End If
    
    MoveToLegalPos = True
End Function

Function InMapBounds(ByVal X As Integer, ByVal Y As Integer) As Boolean
'*****************************************************************
'Checks to see if a tile position is in the maps bounds
'*****************************************************************
    If X < XMinMapSize Or X > XMaxMapSize Or Y < YMinMapSize Or Y > YMaxMapSize Then
        Exit Function
    End If
    
    InMapBounds = True
End Function

Private Sub DDrawGrhtoSurface(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, Color() As Long)
    Dim CurrentGrhIndex As Integer
'On Error GoTo error
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        'Draw
        Device_Textured_Render X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color
        
    End With
Exit Sub

error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub

Sub DDrawTransGrhIndextoSurface(ByVal GrhIndex As Integer, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, Color() As Long)

    With GrhData(GrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If
        
        'Draw
        Device_Textured_Render X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color
    End With
End Sub

Sub DDrawTransGrhtoSurface(ByRef Grh As Grh, ByVal X As Integer, ByVal Y As Integer, ByVal Center As Byte, ByVal Animate As Byte, Color() As Long)
'*****************************************************************
'Draws a GRH transparently to a X and Y position
'*****************************************************************
    Dim CurrentGrhIndex As Integer
    
'On Error GoTo error
    
    If Animate Then
        If Grh.Started = 1 Then
            Grh.FrameCounter = Grh.FrameCounter + (timerElapsedTime * GrhData(Grh.GrhIndex).NumFrames / Grh.Speed)
            
            If Grh.FrameCounter > GrhData(Grh.GrhIndex).NumFrames Then
                Grh.FrameCounter = (Grh.FrameCounter Mod GrhData(Grh.GrhIndex).NumFrames) + 1
                
                If Grh.Loops <> INFINITE_LOOPS Then
                    If Grh.Loops > 0 Then
                        Grh.Loops = Grh.Loops - 1
                    Else
                        Grh.Started = 0
                    End If
                End If
            End If
        End If
    End If
    
    'Figure out what frame to draw (always 1 if not animated)
    CurrentGrhIndex = GrhData(Grh.GrhIndex).Frames(Grh.FrameCounter)
    
    With GrhData(CurrentGrhIndex)
        'Center Grh over X,Y pos
        If Center Then
            If .TileWidth <> 1 Then
                X = X - Int(.TileWidth * TilePixelWidth / 2) + TilePixelWidth \ 2
            End If
            
            If .TileHeight <> 1 Then
                Y = Y - Int(.TileHeight * TilePixelHeight) + TilePixelHeight
            End If
        End If

        'Draw
        Device_Textured_Render X, Y, .pixelWidth, .pixelHeight, .sX, .sY, .FileNum, Color
    End With
Exit Sub

error:
    If Err.Number = 9 And Grh.FrameCounter < 1 Then
        Grh.FrameCounter = 1
        Resume
    Else
        MsgBox "Ocurrió un error inesperado, por favor comuniquelo a los administradores del juego." & vbCrLf & "Descripción del error: " & _
        vbCrLf & Err.Description, vbExclamation, "[ " & Err.Number & " ] Error"
        End
    End If
End Sub

Sub RenderScreen(ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 8/14/2007
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Renders everything to the viewport
'**************************************************************
    Dim Y           As Long     'Keeps track of where on map we are
    Dim X           As Long     'Keeps track of where on map we are
    Dim ScreenMinY  As Integer  'Start Y pos on current screen
    Dim ScreenMaxY  As Integer  'End Y pos on current screen
    Dim ScreenMinX  As Integer  'Start X pos on current screen
    Dim ScreenMaxX  As Integer  'End X pos on current screen
    Dim minY        As Integer  'Start Y pos on current map
    Dim maxY        As Integer  'End Y pos on current map
    Dim minX        As Integer  'Start X pos on current map
    Dim maxX        As Integer  'End X pos on current map
    Dim screenX     As Integer  'Keeps track of where to place tile on screen
    Dim screenY     As Integer  'Keeps track of where to place tile on screen
    Dim minXOffset  As Integer
    Dim minYOffset  As Integer
    Dim PixelOffsetXTemp As Integer 'For centering grhs
    Dim PixelOffsetYTemp As Integer 'For centering grhs
    
    'Figure out Ends and Starts of screen
    ScreenMinY = (UserPos.Y - AddtoUserPos.Y) - HalfWindowTileHeight
    ScreenMaxY = (UserPos.Y - AddtoUserPos.Y) + HalfWindowTileHeight
    ScreenMinX = (UserPos.X - AddtoUserPos.X) - HalfWindowTileWidth
    ScreenMaxX = (UserPos.X - AddtoUserPos.X) + HalfWindowTileWidth
    
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
        screenY = 1
    End If
    
    If ScreenMaxY < YMaxMapSize Then ScreenMaxY = ScreenMaxY + 1
    
    If ScreenMinX > XMinMapSize Then
        ScreenMinX = ScreenMinX - 1
    Else
        ScreenMinX = 1
        screenX = 1
    End If
    
    If ScreenMaxX < XMaxMapSize Then ScreenMaxX = ScreenMaxX + 1

    screenX = screenX - 1
    screenY = screenY - 1
    
    PixelOffsetXTemp = screenX * TilePixelWidth + PixelOffsetX
    
    'Draw floor layer
    For X = ScreenMinX To ScreenMaxX
        For Y = ScreenMinY To ScreenMaxY
                    
            PixelOffsetYTemp = screenY * TilePixelHeight + PixelOffsetY
            
            'Layer 1 **********************************
            Call DDrawGrhtoSurface(MapData(X, Y).Graphic(1), _
                PixelOffsetXTemp, _
                PixelOffsetYTemp, _
                0, 1, AmbientColor)
            '******************************************
            
            'Layer 2 **********************************
            If MapData(X, Y).Graphic(2).GrhIndex <> 0 Then
                Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(2), _
                        PixelOffsetXTemp, _
                        PixelOffsetYTemp, _
                        1, 1, AmbientColor)
            End If
            '******************************************
            
            screenY = screenY + 1
        Next Y
    
        'Reset ScreenX to original value and increment ScreenY
        screenY = screenY - Y + ScreenMinY
        screenX = screenX + 1
        PixelOffsetXTemp = screenX * TilePixelWidth + PixelOffsetX
    Next X
    
   'Draw Transparent Layers
   screenY = (minYOffset - TileBufferSize)
   screenX = (minXOffset - TileBufferSize)
   
   PixelOffsetXTemp = screenX * TilePixelWidth + PixelOffsetX
   
    For X = minX To maxX
        For Y = minY To maxY
        
            PixelOffsetYTemp = screenY * TilePixelHeight + PixelOffsetY
    
            With MapData(X, Y)
                'Object Layer **********************************
                If .ObjGrh.GrhIndex <> 0 Then
                    Call DDrawTransGrhtoSurface(.ObjGrh, _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, AmbientColor)
                End If
                '***********************************************
    
    
                'Char layer ************************************
                If .CharIndex <> 0 Then
                    Call CharRender(.CharIndex, PixelOffsetXTemp, PixelOffsetYTemp)
                End If
                '*************************************************
    
    
                'Layer 3 *****************************************
                If .Graphic(3).GrhIndex <> 0 Then
                    'Draw
                    Call DDrawTransGrhtoSurface(.Graphic(3), _
                            PixelOffsetXTemp, PixelOffsetYTemp, 1, 1, AmbientColor)
                End If
                '************************************************
            End With
    
            screenY = screenY + 1
        Next Y
        screenY = screenY - Y + minY
        screenX = screenX + 1
        PixelOffsetXTemp = screenX * TilePixelWidth + PixelOffsetX
    Next X
        
    If Not bTecho Then
        'Draw blocked tiles and grid
        screenY = (minYOffset - TileBufferSize)
        screenX = (minXOffset - TileBufferSize)
        
        PixelOffsetXTemp = screenX * TilePixelWidth + PixelOffsetX
        
        For X = minX To maxX
            For Y = minY To maxY
               
               PixelOffsetYTemp = screenY * TilePixelHeight + PixelOffsetY
               
                'Layer 4 **********************************
                If MapData(X, Y).Graphic(4).GrhIndex Then
                    'Draw
                    Call DDrawTransGrhtoSurface(MapData(X, Y).Graphic(4), _
                        PixelOffsetXTemp, _
                        PixelOffsetYTemp, _
                        1, 1, AmbientColor)
                End If
                '**********************************
               
                screenY = screenY + 1
            Next Y
            screenY = screenY - Y + minY
            screenX = screenX + 1
            PixelOffsetXTemp = screenX * TilePixelWidth + PixelOffsetX
        
        Next X
    End If
    
End Sub

Public Function RenderSounds()
'**************************************************************
'Author: Juan Martín Sotuyo Dodero
'Last Modify Date: 3/30/2008
'Actualiza todos los sonidos del mapa.
'**************************************************************
    If bLluvia(UserMap) = 1 Then
        If bRain Then
            If bTecho Then
                If frmMain.IsPlaying <> PlayLoop.plLluviain Then
                    'If RainBufferIndex Then _
                        'call 'audio.StopWave(RainBufferIndex)
                    'RainBufferIndex = 'audio.PlayWave("lluviain.wav", 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviain
                End If
            Else
                If frmMain.IsPlaying <> PlayLoop.plLluviaout Then
                    'If RainBufferIndex Then _
                        'call 'audio.StopWave(RainBufferIndex)
                    'RainBufferIndex = 'audio.PlayWave("lluviaout.wav", 0, 0, LoopStyle.Enabled)
                    frmMain.IsPlaying = PlayLoop.plLluviaout
                End If
            End If
        End If
    End If
    
    DoFogataFx
End Function

Function HayUserAbajo(ByVal X As Integer, ByVal Y As Integer, ByVal GrhIndex As Integer) As Boolean
    If GrhIndex > 0 Then
        HayUserAbajo = _
            charlist(UserCharIndex).Pos.X >= X - (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.X <= X + (GrhData(GrhIndex).TileWidth \ 2) _
                And charlist(UserCharIndex).Pos.Y >= Y - (GrhData(GrhIndex).TileHeight - 1) _
                And charlist(UserCharIndex).Pos.Y <= Y
    End If
End Function

Public Function InitTileEngine(ByVal setDisplayFormhWnd As Long, ByVal setMainViewTop As Integer, ByVal setMainViewLeft As Integer, ByVal setTilePixelHeight As Integer, ByVal setTilePixelWidth As Integer, ByVal setWindowTileHeight As Integer, ByVal setWindowTileWidth As Integer, ByVal setTileBufferSize As Integer, ByVal pixelsToScrollPerFrameX As Integer, pixelsToScrollPerFrameY As Integer, ByVal engineSpeed As Single) As Boolean
'***************************************************
'Author: Aaron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Creates all DX objects and configures the engine to start running.
'***************************************************
    IniPath = App.path & "\Init\"
    
    'Fill startup variables
    MainViewTop = setMainViewTop
    MainViewLeft = setMainViewLeft
    TilePixelWidth = setTilePixelWidth
    TilePixelHeight = setTilePixelHeight
    WindowTileHeight = setWindowTileHeight
    WindowTileWidth = setWindowTileWidth
    TileBufferSize = setTileBufferSize
    
    HalfWindowTileHeight = setWindowTileHeight \ 2
    HalfWindowTileWidth = setWindowTileWidth \ 2
    
    'Compute offset in pixels when rendering tile buffer.
    'We diminish by one to get the top-left corner of the tile for rendering.
    'TileBufferPixelOffsetX = ((TileBufferSize - 1) * TilePixelWidth)
    'TileBufferPixelOffsetY = ((TileBufferSize - 1) * TilePixelHeight)
    
    engineBaseSpeed = engineSpeed
    
    'Set FPS value to 60 for startup
    FPS = 60
    FramesPerSecCounter = 60
    
    MinXBorder = XMinMapSize + (WindowTileWidth \ 2)
    MaxXBorder = XMaxMapSize - (WindowTileWidth \ 2)
    MinYBorder = YMinMapSize + (WindowTileHeight \ 2)
    MaxYBorder = YMaxMapSize - (WindowTileHeight \ 2)
    
    MainViewWidth = TilePixelWidth * WindowTileWidth
    MainViewHeight = TilePixelHeight * WindowTileHeight
    
    'Resize mapdata array
    ReDim MapData(XMinMapSize To XMaxMapSize, YMinMapSize To YMaxMapSize) As MapBlock
    
    'Set intial user position
    UserPos.X = MinXBorder
    UserPos.Y = MinYBorder
    
    'Set scroll pixels per frame
    ScrollPixelsPerFrameX = pixelsToScrollPerFrameX
    ScrollPixelsPerFrameY = pixelsToScrollPerFrameY
    
    Call DirectX_Init

    Call LoadGrhData
    Call CargarCuerpos
    Call CargarCabezas
    Call CargarCascos
    Call CargarFxs
    
    Call Engine_Init_FontSettings
    Call Engine_Init_FontTextures
    
    AmbientColor(0) = -1
    AmbientColor(1) = -1
    AmbientColor(2) = -1
    AmbientColor(3) = -1

    InitTileEngine = True
End Function

Public Sub DirectX_Init()

    Dim DispMode As D3DDISPLAYMODE
    Dim PresentationParameters As D3DPRESENT_PARAMETERS
    
    Set DirectX = New DirectX8
    Set DirectD3D = DirectX.Direct3DCreate
    Set DirectD3D8 = New D3DX8
    
    DirectD3D.GetAdapterDisplayMode D3DADAPTER_DEFAULT, DispMode
    
    With PresentationParameters
    
        .Windowed = True
        .SwapEffect = D3DSWAPEFFECT_DISCARD
        
        .BackBufferFormat = DispMode.Format
        .BackBufferWidth = frmMain.MainViewPic.ScaleWidth
        .BackBufferHeight = frmMain.MainViewPic.ScaleHeight
        
    End With

    Set DirectDevice = DirectD3D.CreateDevice(D3DADAPTER_DEFAULT, D3DDEVTYPE_HAL, frmMain.MainViewPic.hWnd, D3DCREATE_HARDWARE_VERTEXPROCESSING, PresentationParameters)
    
    Call D3DXMatrixOrthoOffCenterLH(Projection, 0, frmMain.MainViewPic.ScaleWidth, frmMain.MainViewPic.ScaleHeight, 0, -1#, 1#)
    Call D3DXMatrixIdentity(View)
    
    Call DirectDevice.SetTransform(D3DTS_PROJECTION, Projection)
    Call DirectDevice.SetTransform(D3DTS_VIEW, View)
    
    Engine_Init_RenderStates
    
    Set SurfaceDB = New clsSurfaceDB
    Set SpriteBatch = New clsBatch
    
    Call SurfaceDB.Init(DirectD3D8, DirectDevice, 90)
    Call SpriteBatch.Initialise(1000)
    
    If DirectDevice Is Nothing Then
        MsgBox "No se puede inicializar DirectX. Por favor asegúrese de tener la última versión correctamente instalada."
        Exit Sub
    End If
    
End Sub
Private Sub Engine_Init_RenderStates()

    DirectDevice.SetVertexShader D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1
        
    'Set the render states
    With DirectDevice
    
        .SetRenderState D3DRS_LIGHTING, False
        .SetRenderState D3DRS_SRCBLEND, D3DBLEND_SRCALPHA
        .SetRenderState D3DRS_DESTBLEND, D3DBLEND_INVSRCALPHA
        .SetRenderState D3DRS_ALPHABLENDENABLE, True
        .SetRenderState D3DRS_FILLMODE, D3DFILL_SOLID
        .SetRenderState D3DRS_CULLMODE, D3DCULL_NONE
        .SetTextureStageState 0, D3DTSS_ALPHAOP, D3DTOP_MODULATE
    
    End With
    
End Sub

Public Sub DeinitTileEngine()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modification: 08/14/07
'Destroys all DX objects
'***************************************************
'Set no texture in the device to avoid memory leaks

        If Not DirectDevice Is Nothing Then
                DirectDevice.SetTexture 0, Nothing
        End If
        
        '// Destroy Textures
        Set SurfaceDB = Nothing
        
        Set DirectX = Nothing
        Set DirectD3D = Nothing
        Set DirectDevice = Nothing
        
        Set SpriteBatch = Nothing
        
        'Clear arrays
        Erase GrhData
        Erase BodyData
        Erase HeadData
        Erase FxData
        Erase WeaponAnimData
        Erase ShieldAnimData
        Erase CascoAnimData
        Erase MapData
        Erase charlist
        
End Sub


Public Sub ShowNextFrame(ByVal DisplayFormTop As Integer, ByVal DisplayFormLeft As Integer, ByVal MouseViewX As Integer, ByVal MouseViewY As Integer)
'***************************************************
'Author: Arron Perkins
'Last Modification: 08/14/07
'Last modified by: Juan Martín Sotuyo Dodero (Maraxus)
'Updates the game's model and renders everything.
'***************************************************
    Static OffsetCounterX As Single
    Static OffsetCounterY As Single
    
    If EngineRun Then
        If UserMoving Then
            '****** Move screen Left and Right if needed ******
            If AddtoUserPos.X <> 0 Then
                OffsetCounterX = OffsetCounterX - ScrollPixelsPerFrameX * AddtoUserPos.X * timerTicksPerFrame
                If Abs(OffsetCounterX) >= Abs(TilePixelWidth * AddtoUserPos.X) Then
                    OffsetCounterX = 0
                    AddtoUserPos.X = 0
                    UserMoving = False
                End If
            End If
            
            '****** Move screen Up and Down if needed ******
            If AddtoUserPos.Y <> 0 Then
                OffsetCounterY = OffsetCounterY - ScrollPixelsPerFrameY * AddtoUserPos.Y * timerTicksPerFrame
                If Abs(OffsetCounterY) >= Abs(TilePixelHeight * AddtoUserPos.Y) Then
                    OffsetCounterY = 0
                    AddtoUserPos.Y = 0
                    UserMoving = False
                End If
            End If
        End If
        
        'Update mouse position within view area
        Call ConvertCPtoTP(MouseViewX, MouseViewY, MouseTileX, MouseTileY)
        
        'DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET, 0, 1#, 0
        DirectDevice.BeginScene
        
        Call SpriteBatch.Begin
        
        '****** Update screen ******
        If UserCiego Then
            Call CleanViewPort
        Else
            Call RenderScreen(OffsetCounterX, OffsetCounterY)
        End If
        
        Call Dialogos.Render(SpriteBatch)
        
        Call DibujarCartel

       ' Call DialogosClanes.Draw
        
        Call SpriteBatch.Finish
        
        DirectDevice.EndScene
        DirectDevice.Present ByVal 0, ByVal 0, 0, ByVal 0
        
        'FPS update
        If fpsLastCheck + 1000 < GetTickCount() Then
            FPS = FramesPerSecCounter
            FramesPerSecCounter = 1
            fpsLastCheck = GetTickCount()
        Else
            FramesPerSecCounter = FramesPerSecCounter + 1
        End If
        
        'Get timing info
        timerElapsedTime = GetElapsedTime()
        timerTicksPerFrame = timerElapsedTime * engineBaseSpeed
    End If
End Sub

Private Function GetElapsedTime() As Single
'**************************************************************
'Author: Aaron Perkins
'Last Modify Date: 10/07/2002
'Gets the time that past since the last call
'**************************************************************
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

Private Sub CharRender(ByVal CharIndex As Long, ByVal PixelOffsetX As Integer, ByVal PixelOffsetY As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Draw char's to screen without offcentering them
'***************************************************
    Dim moved As Boolean
    Dim Pos As Integer
    Dim line As String
    Dim Color(3) As Long
    
With charlist(CharIndex)
        If .Moving Then
            'If needed, move left and right
            If .scrollDirectionX <> 0 Then
                .MoveOffsetX = .MoveOffsetX + ScrollPixelsPerFrameX * Sgn(.scrollDirectionX) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionX) = 1 And .MoveOffsetX >= 0) Or (Sgn(.scrollDirectionX) = -1 And .MoveOffsetX <= 0) Then
                    .MoveOffsetX = 0
                    .scrollDirectionX = 0
                End If
            End If
            
            'If needed, move up and down
            If .scrollDirectionY <> 0 Then
                .MoveOffsetY = .MoveOffsetY + ScrollPixelsPerFrameY * Sgn(.scrollDirectionY) * timerTicksPerFrame
                
                'Start animations
'TODO : Este parche es para evita los uncornos exploten al moverse!! REVER!!!
                If .Body.Walk(.Heading).Speed > 0 Then _
                    .Body.Walk(.Heading).Started = 1
                .Arma.WeaponWalk(.Heading).Started = 1
                .Escudo.ShieldWalk(.Heading).Started = 1
                
                'Char moved
                moved = True
                
                'Check if we already got there
                If (Sgn(.scrollDirectionY) = 1 And .MoveOffsetY >= 0) Or (Sgn(.scrollDirectionY) = -1 And .MoveOffsetY <= 0) Then
                    .MoveOffsetY = 0
                    .scrollDirectionY = 0
                End If
            End If
        End If
        
        'If done moving stop animation
        If Not moved Then
            'Stop animations
            .Body.Walk(.Heading).Started = 0
            .Body.Walk(.Heading).FrameCounter = 1
            
            .Arma.WeaponWalk(.Heading).Started = 0
            .Arma.WeaponWalk(.Heading).FrameCounter = 1
            
            .Escudo.ShieldWalk(.Heading).Started = 0
            .Escudo.ShieldWalk(.Heading).FrameCounter = 1
            
            .Moving = False
        End If
        
        PixelOffsetX = PixelOffsetX + .MoveOffsetX
        PixelOffsetY = PixelOffsetY + .MoveOffsetY
        
        If .Head.Head(.Heading).GrhIndex Then
            If Not .invisible Then
                'Draw Body
                If .Body.Walk(.Heading).GrhIndex Then Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, AmbientColor)
            
                'Draw Head
                If .Head.Head(.Heading).GrhIndex Then Call DDrawTransGrhtoSurface(.Head.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, AmbientColor)
                    
                'Draw Helmet
                If .Casco.Head(.Heading).GrhIndex Then Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, 1, 0, AmbientColor)
                    ' Call DDrawTransGrhtoSurface(.Casco.Head(.Heading), PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, 1, 0)
                
                'Draw Weapon
                If .Arma.WeaponWalk(.Heading).GrhIndex Then Call DDrawTransGrhtoSurface(.Arma.WeaponWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, AmbientColor)
                    
                'Draw Shield
                If .Escudo.ShieldWalk(.Heading).GrhIndex Then Call DDrawTransGrhtoSurface(.Escudo.ShieldWalk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, AmbientColor)
                
                If LenB(.Nombre) > 0 Then
                    If Nombres Then
                        Pos = getTagPosition(.Nombre)
                            'Pos = InStr(.Nombre, "<")
                            'If Pos = 0 Then Pos = Len(.Nombre) + 2

                        If .priv = 0 Then
                            Select Case .Criminal
                            
                                Case 2 'ciuda
                                    Color(0) = ColoresPJ(49)
                                    Color(1) = ColoresPJ(49)
                                    Color(2) = ColoresPJ(49)
                                    Color(3) = ColoresPJ(49)
                                Case 3
                                    Color(0) = ColoresPJ(50)
                                    Color(1) = ColoresPJ(50)
                                    Color(2) = ColoresPJ(50)
                                    Color(3) = ColoresPJ(50)
                                Case 4
                                    Color(0) = ColoresPJ(47)
                                    Color(1) = ColoresPJ(47)
                                    Color(2) = ColoresPJ(47)
                                    Color(3) = ColoresPJ(47)
                                Case 5
                                    Color(0) = ColoresPJ(48)
                                    Color(1) = ColoresPJ(48)
                                    Color(2) = ColoresPJ(48)
                                    Color(3) = ColoresPJ(48)
                            End Select
                        Else
                            Color(0) = ColoresPJ(.priv)
                            Color(1) = ColoresPJ(.priv)
                            Color(2) = ColoresPJ(.priv)
                            Color(3) = ColoresPJ(.priv)
                        End If
                            
                            'Nick
                            line = Left$(.Nombre, Pos - 2)
                            Call Text_Draw(SpriteBatch, PixelOffsetX - .NombreOffset, PixelOffsetY + 30, line, Color)
                            
                            'Clan
                         '   line = mid$(.Nombre, Pos)
                           ' Call RenderTextCentered(PixelOffsetX + TilePixelWidth \ 2 + 5, PixelOffsetY + 45, line, color, frmMain.font)
                    End If
                End If
            End If
        Else
            'Draw Body
            If .Body.Walk(.Heading).GrhIndex Then _
                Call DDrawTransGrhtoSurface(.Body.Walk(.Heading), PixelOffsetX, PixelOffsetY, 1, 1, AmbientColor)
        End If
        
        'Update dialogs
        'Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y + OFFSET_HEAD, CharIndex) '34 son los pixeles del grh de la cabeza que quedan superpuestos al cuerpo
        Call Dialogos.UpdateDialogPos(PixelOffsetX + .Body.HeadOffset.X, PixelOffsetY + .Body.HeadOffset.Y, CharIndex)
        
        'Draw FX
        If .FxIndex <> 0 Then
            'Check if animation is over
            If .fX.Started = 0 Then .FxIndex = 0
        End If
        
End With

End Sub

Public Sub SetCharacterFx(ByVal CharIndex As Integer, ByVal fX As Integer, ByVal Loops As Integer)
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Sets an FX to the character.
'***************************************************
    With charlist(CharIndex)
        .FxIndex = fX
        
        If .FxIndex > 0 Then
            Call InitGrh(.fX, FxData(fX).Animacion)
        
            .fX.Loops = Loops
        End If
    End With
End Sub

Private Sub CleanViewPort()
'***************************************************
'Author: Juan Martín Sotuyo Dodero (Maraxus)
'Last Modify Date: 12/03/04
'Fills the viewport with black.
'***************************************************
    'todo: check if inventory still rendering after this
    DirectDevice.Clear 0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, 0, 1#, 0
End Sub

Public Sub Device_Textured_Render(ByVal X As Integer, ByVal Y As Integer, _
                                  ByVal Width As Integer, ByVal Height As Integer, _
                                  ByVal sX As Integer, ByVal sY As Integer, _
                                  tex As Long, _
                                  ByRef Color() As Long)
        
        Dim Texture As Direct3DTexture8
        Dim TexWidth As Long, TexHeight As Long
        
        Set Texture = SurfaceDB.GetTexture(tex, TexWidth, TexHeight)

        With SpriteBatch
                '// Seteamos la textura
                Call .SetTexture(Texture)
                
                If TexWidth <> 0 And TexHeight <> 0 Then
                    Call .Draw(X, Y, Width, Height, Color, sX / TexWidth, sY / TexHeight, (sX + Width) / TexWidth, (sY + Height) / TexHeight)
                Else
                    Call .Draw(X, Y, TexWidth, TexHeight, Color)
                End If
        End With

End Sub

Public Sub SetCamera(ByVal X As Single, ByVal Y As Single)
    
    'Camera.X = X
    'Camera.Y = Y
    
    'With At
    
    '    .X = Camera.X
    '    .Y = Camera.Y
    '    .Z = 0
        
    'End With
    
    'Call D3DXMATH_MATRIX.D3DXMatrixLookAtLH(View, Camera, At, Up)
    'Call DirectDevice.SetTransform(D3DTS_VIEW, View)
End Sub

Public Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
    Static LastMovement As Long
    
    'No input allowed while Argentum is not the active window
    If Not Application.IsAppActive() Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'No walking while writting in the forum.
    If MirandoForo Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    'TODO: Debería informarle por consola?
    If Traveling Then Exit Sub

    'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
    If GetTickCount - LastMovement > 56 Then
        LastMovement = GetTickCount
    Else
        Exit Sub
    End If
    
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(NORTH)
                'todo
                frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
                Exit Sub
            End If
            
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(EAST)
                frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.X & "," & UserPos.Y & ")"
                frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(SOUTH)
                frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(WEST)
                frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
                Exit Sub
            End If
            
            ' We haven't moved - Update 3D sounds!
            'call 'audio.MoveListener(UserPos.X, UserPos.Y)
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call RandomMove
            Else
                ' We haven't moved - Update 3D sounds!
                'call 'audio.MoveListener(UserPos.X, UserPos.Y)
            End If
            
            If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
            frmMain.Coord.Caption = "(" & UserPos.X & "," & UserPos.Y & ")"
            frmMain.Coord.Caption = "X: " & UserPos.X & " Y: " & UserPos.Y
        End If
    End If
End Sub
