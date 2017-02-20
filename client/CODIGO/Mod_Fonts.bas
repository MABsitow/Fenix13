Attribute VB_Name = "Mod_Fonts"
 Option Explicit

Private Type D3DXIMAGE_INFO_A
    Width As Long
    Height As Long
    Depth As Long
    MipLevels As Long
    Format As CONST_D3DFORMAT
    ResourceType As CONST_D3DRESOURCETYPE
    ImageFileFormat As Long
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
    CharWidth(0 To 255) As Byte 'The actual factual width of each character
    CharVA(0 To 255) As CharVA
End Type

Private Type CustomFont
    HeaderInfo As VFH           'Holds the header information
    Texture As Direct3DTexture8 'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As POINTAPI     'Size of the texture
End Type

Public cfonts(1 To 2) As CustomFont ' _Default2 As CustomFont


'ignora esto
Public Sub Text_Draw(ByVal Left As Long, ByVal Top As Long, ByVal Text As String, ByVal Color As Long, Optional ByVal Alpha As Byte = 255, Optional ByVal Center As Boolean = False, Optional ByVal fontNum As Byte = 1)
    
    If Alpha <> 255 Then
        Dim aux As D3DCOLORVALUE
        
        'Obtener_RGB Color, r, g, b
        ARGBtoD3DCOLORVALUE Color, aux
        
        Color = D3DColorARGB(Alpha, aux.r, aux.g, aux.b)
    End If

    Engine_Render_Text cfonts(fontNum), Text, Left, Top, Color, Alpha

End Sub

Private Sub Engine_Render_Text(ByVal Batch As clsBGFXSpriteBatch, ByRef UseFont As CustomFont, ByVal Text As String, ByVal X As Long, ByVal Y As Long, ByVal Color As Long, Optional ByVal Alpha As Byte = 255)
'*****************************************************************
'Render text with a custom font
'*****************************************************************
    Dim TempVA(0 To 3) As TLVERTEX
    Dim tempstr() As String
    Dim Count As Integer
    Dim ascii() As Byte
    Dim i As Long
    Dim j As Long
    Dim TempColor As Long
    Dim ResetColor As Byte
    Dim YOffset As Single

    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
    
    'tempstr = Split(Text, Chr(32))
    
    'Get the text into arrays (split by vbCrLf)
    tempstr = Split(Text, vbCrLf)
    
    'Set the temp color (or else the first character has no color)
    TempColor = Color

    'Set the texture
    Batch.SetTexture UseFont.Texture
    
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
                CopyMemory TempVA(0), UseFont.HeaderInfo.CharVA(ascii(j - 1)).Vertex(0), FVF_SIZE * 4
                
                'Set up the verticies
                TempVA(0).X = X + Count
                TempVA(0).Y = Y + YOffset
                
                TempVA(1).X = TempVA(1).X + X + Count
                TempVA(1).Y = TempVA(0).Y
                
                TempVA(2).X = TempVA(0).X
                TempVA(2).Y = TempVA(2).Y + TempVA(0).Y
                
                TempVA(3).X = TempVA(1).X
                TempVA(3).Y = TempVA(2).Y
                
                TempVA(0).Color = TempColor
                TempVA(1).Color = TempColor
                TempVA(2).Color = TempColor
                TempVA(3).Color = TempColor
            
                Batch.Draw X + Count, Y + YOffset, UseFont.HeaderInfo.CellWidth, UseFont.HeaderInfo.CellHeight, Color, _
                    TempVA(0).tu, TempVA(0).tv, TempVA(3).tu, TempVA(3).tv
                
                'Shift over the the position to render the next character
                Count = Count + UseFont.HeaderInfo.CharWidth(ascii(j - 1))
                
                'Check to reset the color
                If ResetColor Then
                    ResetColor = 0
                    TempColor = Color
                End If
                
            Next j
            
        End If
    Next i
    
End Sub

Private Function ARGBtoD3DCOLORVALUE(ByVal ARGB As Long, ByRef Color As D3DCOLORVALUE)
Dim dest(3) As Byte
CopyMemory dest(0), ARGB, 4
Color.a = dest(3)
Color.r = dest(2)
Color.g = dest(1)
Color.b = dest(0)
End Function

Public Function Text_GetWidth(ByRef UseFont As CustomFont, ByVal Text As String) As Integer
'***************************************************
'Returns the width of text
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_GetTextWidth
'***************************************************
Dim i As Integer

    'Make sure we have text
    If LenB(Text) = 0 Then Exit Function
    
    'Loop through the text
    For i = 1 To Len(Text)
        
        'Add up the stored character widths
        Text_GetWidth = Text_GetWidth + UseFont.HeaderInfo.CharWidth(Asc(mid$(Text, i, 1)))
        
    Next i

End Function

Sub Engine_Init_FontTextures()
On Error GoTo eDebug:
'*****************************************************************
'Init the custom font textures
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontTextures
'*****************************************************************
    Dim TexInfo As D3DXIMAGE_INFO_A

    'Check if we have the device
    If DirectDevice.TestCooperativeLevel <> D3D_OK Then Exit Sub

    '*** Default font ***
    
    'Set the texture
    Set cfonts(1).Texture = DirectD3DX.CreateTextureFromFileEx(DirectDevice, DirGraficos & "Font.png", D3DX_DEFAULT, D3DX_DEFAULT, 0, 0, D3DFMT_UNKNOWN, D3DPOOL_MANAGED, D3DX_FILTER_POINT, D3DX_FILTER_NONE, 0, TexInfo, ByVal 0)
    
    'Store the size of the texture
    cfonts(1).TextureSize.X = TexInfo.Width
    cfonts(1).TextureSize.Y = TexInfo.Height
    
    Exit Sub
eDebug:
    If Err.Number = "-2005529767" Then
        MsgBox "Error en la textura de fuente utilizada " & DirGraficos & "Font.png.", vbCritical
        End
    End If
    End

End Sub

Sub Engine_Init_FontSettings()
    '*****************************************************************
    'Init the custom font settings
    'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontSettings
    '*****************************************************************
    Dim FileNum  As Byte
    Dim LoopChar As Long
    Dim Row      As Single
    Dim u        As Single
    Dim v        As Single

    '*** Default font ***

    'Load the header information
    FileNum = FreeFile
    Open IniPath & "Font.dat" For Binary As #FileNum
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
        v = Row * cfonts(1).RowFactor

        ' esos 4 necesitas, se, pero tengo que llevar los cálculos de arriba al sub jaja normalmente se crea una struct llamad Glyph que contiene eso mañana le voy a pegar un rework importante jajaj

        'Set the verticies
        With cfonts(1).HeaderInfo.CharVA(LoopChar)
            .Vertex(0).Color = D3DColorARGB(255, 0, 0, 0)   'Black is the most common color
            '.Vertex(0).Rhw = 1
            .Vertex(0).tu = u
            .Vertex(0).tv = v
            .Vertex(0).X = 0
            .Vertex(0).Y = 0
            .Vertex(0).Z = 0
            
            .Vertex(1).Color = D3DColorARGB(255, 0, 0, 0)
            '.Vertex(1).Rhw = 1
            .Vertex(1).tu = u + cfonts(1).ColFactor
            .Vertex(1).tv = v
            .Vertex(1).X = cfonts(1).HeaderInfo.CellWidth
            .Vertex(1).Y = 0
            .Vertex(1).Z = 0
            
            .Vertex(2).Color = D3DColorARGB(255, 0, 0, 0)
            '.Vertex(2).Rhw = 1
            .Vertex(2).tu = u
            .Vertex(2).tv = v + cfonts(1).RowFactor
            .Vertex(2).X = 0
            .Vertex(2).Y = cfonts(1).HeaderInfo.CellHeight
            .Vertex(2).Z = 0
            
            .Vertex(3).Color = D3DColorARGB(255, 0, 0, 0)
            '.Vertex(3).Rhw = 1
            .Vertex(3).tu = u + cfonts(1).ColFactor
            .Vertex(3).tv = v + cfonts(1).RowFactor
            .Vertex(3).X = cfonts(1).HeaderInfo.CellWidth
            .Vertex(3).Y = cfonts(1).HeaderInfo.CellHeight
            .Vertex(3).Z = 0
        End With
        
    Next LoopChar
    
    Call FontBatch.Initialise(1000) ' no
    Call FontBatch.SetTexture(cfonts(1).Texture) 'por ahora no importa que este aca, total es lo unico que vamos a dibujar
    'claro, yo lo decia cuando este todo :p
    'Set FontVB = DirectDevice.CreateVertexBuffer(FVF_SIZE * 1000 * 4, 0, FVF, D3DPOOL_MANAGED)
    
    '
    ' Pre crear indices (wolfy)
    '
    'Set FontIB = DirectDevice.CreateIndexBuffer(1000 * 6, 0, D3DFMT_INDEX16, D3DPOOL_MANAGED)
    
    'Dim lpIndices(0 To 1000 * 6) As Integer

    'Dim i As Long, j As Long
    'For i = 0 To UBound(lpIndices) - 1 Step 6
    '    lpIndices(i) = j
    '    lpIndices(i + 1) = j + 1
    '    lpIndices(i + 2) = j + 2
    '    lpIndices(i + 3) = j + 2
    '    lpIndices(i + 4) = j + 3
    '   lpIndices(i + 5) = j
    '   j = j + 4
    'Next
    ' donde crasheO?
    'despues del setdata de aca si yo tengo Array(1) cuando devuelve UBound? 1
    'Call D3DIndexBuffer8SetData(FontIB, 0, UBound(lpIndices), 0, lpIndices(0))

End Sub
