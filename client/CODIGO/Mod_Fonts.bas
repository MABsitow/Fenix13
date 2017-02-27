Attribute VB_Name = "Mod_Fonts"
 Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef destination As Any, ByRef source As Any, ByVal length As Long)


Private Type POINTAPI
    X As Long
    Y As Long
End Type
        
Private Type CharVA
    X As Integer
    Y As Integer
    W As Integer
    H As Integer
    
    Tx1 As Single
    Tx2 As Single
    Ty1 As Single
    Ty2 As Single
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
    Texture As Long             'Holds the texture of the text
    RowPitch As Integer         'Number of characters per row
    RowFactor As Single         'Percentage of the texture width each character takes
    ColFactor As Single         'Percentage of the texture height each character takes
    CharHeight As Byte          'Height to use for the text - easiest to start with CellHeight value, and keep lowering until you get a good value
    TextureSize As POINTAPI     'Size of the texture
End Type

Public cfonts(1 To 2) As CustomFont ' _Default2 As CustomFont

Public Sub Text_Draw(ByRef Batch As clsBGFXSpriteBatch, ByVal Left As Long, ByVal Top As Long, ByVal Text As String, Color() As Long, Optional ByVal Center As Boolean = False, Optional ByVal fontNum As Byte = 1)

    Engine_Render_Text Batch, cfonts(fontNum), Text, Left, Top, Color

End Sub

Private Sub Engine_Render_Text(ByRef Batch As clsBGFXSpriteBatch, ByRef UseFont As CustomFont, ByVal Text As String, ByVal X As Long, ByVal Y As Long, Color() As Long)
'*****************************************************************
'Render text with a custom font
'*****************************************************************
    Dim TempVA As CharVA
    Dim tempstr() As String
    Dim Count As Integer
    Dim ascii() As Byte
    Dim i As Long
    Dim j As Long
    Dim ResetColor As Byte
    Dim yOffset As Single

    'Check for valid text to render
    If LenB(Text) = 0 Then Exit Sub
    
    'tempstr = Split(Text, Chr(32))
    
    'Get the text into arrays (split by vbCrLf)
    tempstr = Split(Text, vbCrLf)
    

    'Set the texture
    Batch.SetTexture UseFont.Texture
    
    'Loop through each line if there are line breaks (vbCrLf)
    For i = 0 To UBound(tempstr)
        If Len(tempstr(i)) > 0 Then
            yOffset = i * UseFont.CharHeight
            Count = 0
        
            'Convert the characters to the ascii value
            ascii() = StrConv(tempstr(i), vbFromUnicode)
        
            'Loop through the characters
            For j = 1 To Len(tempstr(i))

                'Copy from the cached vertex array to the temp vertex data
                CopyMemory TempVA, UseFont.HeaderInfo.CharVA(ascii(j - 1)), 24 'this number represents the size of "CharVA" struct
                
                TempVA.X = X + Count
                TempVA.Y = Y + yOffset
            
                Batch.Draw TempVA.X, TempVA.Y, TempVA.W, TempVA.H, Color, _
                    TempVA.Tx1, TempVA.Ty1, TempVA.Tx2, TempVA.Ty2
                
                'Shift over the the position to render the next character
                Count = Count + UseFont.HeaderInfo.CharWidth(ascii(j - 1))
                
            Next j
            
        End If
    Next i
    
End Sub

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
'*****************************************************************
'Init the custom font textures
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Init_FontTextures
'*****************************************************************
    Dim Image As TYPE_VIDEO_IMAGE

    '*** Default font ***
    
    Image = Video.CreateImageFromFilename(DirGraficos & "Font.png")
    
    'Set the texture
    cfonts(1).Texture = Image.mHandle
    
    'Store the size of the texture
    cfonts(1).TextureSize.X = Image.mX
    cfonts(1).TextureSize.Y = Image.mY

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

        'Set the verticies
        With cfonts(1).HeaderInfo.CharVA(LoopChar)
            .X = 0
            .Y = 0
            .W = cfonts(1).HeaderInfo.CellWidth
            .H = cfonts(1).HeaderInfo.CellHeight
            .Tx1 = u
            .Ty1 = v
            .Tx2 = u + cfonts(1).ColFactor
            .Ty2 = v + cfonts(1).RowFactor
        End With
        
    Next LoopChar
    
End Sub
