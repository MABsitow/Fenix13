Attribute VB_Name = "Mod_BMFonts"
'--------------------------------------------------------------------------------
'    Component  : Mod_BMFonts
'    Project    : ClientFenix13
'
'    Description: Load the BMFont binary file
'    Author     : Facundo Ortega (GoDKeR)
'--------------------------------------------------------------------------------

Option Explicit

'
'Each block starts with a one byte block type identifier, followed by a 4 byte integer that gives the size of the block,
'not including the block type identifier and the size value.
'
Private Type BLOCK_HEADER
            blockID     As Byte
            blockSize   As Long
End Type


'
'This structure gives the layout of the fields. Remember that there should be no padding between members.
'Allocate the size of the block using the blockSize, as following the block comes the font name, including the
'terminating null char. Most of the time this block can simply be ignored.
'
Private Type BLOCK_TYPE_INFO
            fontSize        As Integer
            bitField        As Byte
            charSet         As Byte
            stretchH        As Integer
            aa              As Byte
            paddingUp       As Byte
            paddingRight    As Byte
            paddingDown     As Byte
            paddingLeft     As Byte
            spacingHoriz    As Byte
            spacingVert     As Byte
            outline         As Byte
            fontName        As String
End Type

Private Type BLOCK_TYPE_COMMON
            lineHeight      As Integer
            base            As Integer
            scaleW          As Integer
            scaleH          As Integer
            pages           As Integer
            bitFiled        As Byte
            alphaChnl       As Byte
            redChnl         As Byte
            greenChnl       As Byte
            blueChnl        As Byte
End Type

'
'This block gives the name of each texture file with the image data for the characters. The string pageNames holds the
'names separated and terminated by null chars. Each filename has the same length, so once you know the size of the first name,
'you can easily determine the position of each of the names. The id of each page is the zero-based index of the string name.
'
Private Type BLOCK_TYPE_PAGES
            pageNames       As String
End Type

'
'The number of characters in the file can be computed by taking the size of the block and dividing with the size of
'the charInfo structure, i.e.: numChars = charsBlock.blockSize/20.
'
Private Type BLOCK_TYPE_CHARS
            id              As Long
            x               As Integer
            y               As Integer
            width           As Integer
            height          As Integer
            xOffset         As Integer
            yOffset         As Integer
            xAdvance        As Integer
            page            As Byte
            chnl            As Byte
End Type

'
'This block is only in the file if there are any kerning pairs with amount differing from 0.
'
Private Type BLOCK_TYPE_KERNING_PAIRS
            first           As Long
            Second          As Long
            amount          As Integer
End Type

Private Type TYPE_FONT
            infoBlock       As BLOCK_TYPE_INFO
            commonBlock     As BLOCK_TYPE_COMMON
            pagesBlock      As BLOCK_TYPE_PAGES
            charsBlock()    As BLOCK_TYPE_CHARS
            kerningBlock()  As BLOCK_TYPE_KERNING_PAIRS
End Type

Public Font As TYPE_FONT

Public Sub LoadFontData()
    Dim handle As Integer
    
    Dim Header As BLOCK_HEADER
    
    Dim Read As New clsByteBuffer
    Dim data() As Byte
    
    handle = FreeFile()
    
    Open App.path & "\INIT\Font\Font.fnt" For Binary Access Read As handle
        ReDim data(0 To LOF(handle) - 1) As Byte
        
        Get handle, , data
    Close handle
    
    With Read
    
        Read.initializeReader data
        
        Call .getVoid(4)
        
        Header.blockID = .getByte
        Header.blockSize = .getLong
        
        With Font.infoBlock
            .fontSize = Read.getInteger
            .bitField = Read.getByte
            .charSet = Read.getByte
            .stretchH = Read.getInteger
            .aa = Read.getByte
            .paddingUp = Read.getByte
            .paddingRight = Read.getByte
            .paddingDown = Read.getByte
            .paddingLeft = Read.getByte
            .spacingHoriz = Read.getByte
            .spacingVert = Read.getByte
            .outline = Read.getByte
            .fontName = Read.getString(Header.blockSize - Read.getCurrentPos)
        End With
        
        Header.blockID = .getByte
        Header.blockSize = .getLong
        
    End With

End Sub
