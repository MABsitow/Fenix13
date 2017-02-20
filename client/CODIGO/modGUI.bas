Attribute VB_Name = "modGUI"
Option Explicit

Public Declare Function ArrPtr Lib "msvbvm60" Alias "VarPtr" (Ptr() As Any) As Long

Public Enum eGuiObjectType
    Label
    TextBox
    Button
End Enum

Public Enum eGuiElement
    TopLeft
    TopRight
    TopMid
    BotLeft
    BotRight
    BotMid
    ExitBtn
    BG
    LeftMid
    RightMid
    RightButton
    MidButton
    LeftButton
    Effect
End Enum

Private Type Texture
        'Tex As Direct3DTexture8
        Width As Integer
        Height As Integer
End Type

'atlas
'Private GUITex As Texture

'files
Private Type tGUIImages
        ID As Long
        X As Long
        Y As Long
        W As Long
        H As Long
End Type

Private Image() As tGUIImages

Private frm As New clsGUIWindow
'Private WindowVB As Direct3DVertexBuffer8
'Private WindowVerts() As TLVERTEX

'Private Verts() As TLVERTEX

'Private GUIBatch As New clsBatch

Public Sub LoadGUIImage()
    Dim handle As Integer
    Dim elements As Integer
    Dim surfaceDesc As D3DSURFACE_DESC
    
    handle = FreeFile
    
    Open IniPath & "gui.fnx" For Binary As handle
        Seek handle, 1
        
        elements = NumbersOfElements(LOF(handle))
        
        ReDim Image(0 To elements - 1) As tGUIImages
        
        Get handle, , Image
        
    Close handle

    Set GUITex.Tex = DirectD3DX.CreateTextureFromFile(DirectDevice, DirGraficos & "Gui.png")
    
    Call GUITex.Tex.GetLevelDesc(0, surfaceDesc)
    
    GUITex.Width = surfaceDesc.Width
    GUITex.Height = surfaceDesc.Height
    
    ReDim Verts(0 To 6 * elements) As TLVERTEX
    Call FillVerts
    
    Call frm.Initialize(50, 50, 320, 320)
    
    Dim events() As Long
    
    Call frm.AddObject("Lbl1", 25, 25, 0, 0, eGuiObjectType.Label, events, "Version: FXIII")
    Call frm.AddObject("Btn1", 50, 50, 96, 53, eGuiObjectType.Button, events, "Hola")
    Call frm.AddObject("Lbl2", 50, 300, 0, 0, eGuiObjectType.Label, events, "LL", True)

    Dim Count As Integer
    Count = frm.getVertsCount
    
    ReDim WindowVerts(0 To Count) As TLVERTEX
    
    Call frm.getWindowVerts(WindowVerts)
    
    GUIBatch.Initialise Count
    
    'Set WindowVB = DirectDevice.CreateVertexBuffer(FVF_SIZE * frm.getElementsCount * 6, 0, FVF, D3DPOOL_MANAGED)
    
    'D3DVertexBuffer8SetData WindowVB, 0, FVF_SIZE * frm.getElementsCount * 6, 0, WindowVerts(0)
    
End Sub

Public Sub DrawAllWindows()
    
    'DirectDevice.SetTexture 0, GUITex.Tex
    
    'DirectDevice.SetStreamSource 0, WindowVB, FVF_SIZE
    
    'DirectDevice.DrawPrimitive D3DPT_TRIANGLELIST, 0, frm.getElementsCount * 2
    
    GUIBatch.Begin
    
    Call frm.OnDraw(GUIBatch)
    
    GUIBatch.Finish
    
    'Text_Render
    'Dim i As Long
    
    'For i = 1 To frm.getObjectCount
    '    Text_Draw frm.getObject(i).ObjectRectangle.x1, frm.getObject(i).ObjectRectangle.Y1, frm.getObject(i).Text, -1
    'Next
    
End Sub

Public Sub GuiMouseMove(ByVal X As Integer, ByVal Y As Integer)
'te extraño for each

Dim I As Long

'For i = 0 To windowsCount
    
    I = frm.Collide(X, Y)
    If I <> -1 Then
    
        'Call Text_Colorize(frm.getObject(I).TSV, Len(frm.getObject(I).Text), D3DColorXRGB(255, 0, 0))
        
    End If
    
    
End Sub

Public Sub VertsFromElement(ByVal ID As Integer, ByRef outVerts() As TLVERTEX, Optional ByVal StartIndex As Integer = 0) ', Optional ByRef vertsCount As Long = -1)
    CopyMemory outVerts(StartIndex), Verts(ID * 6), FVF_SIZE * 6
    
    'If vertsCount <> -1 Then
    '    vertsCount = vertsCount + 6
    'End If
    
End Sub

Public Sub Translate(ByRef outVerts() As TLVERTEX, ByVal StartIndex As Integer, ByVal X As Integer, ByVal Y As Integer)
    
    outVerts(StartIndex).X = X
    outVerts(StartIndex).Y = Y
    
    outVerts(StartIndex + 1).X = outVerts(StartIndex + 1).X + X
    outVerts(StartIndex + 1).Y = Y
    
    outVerts(StartIndex + 2).X = X
    outVerts(StartIndex + 2).Y = outVerts(StartIndex + 2).Y + Y
    
    outVerts(StartIndex + 4).X = outVerts(StartIndex + 4).X + X
    outVerts(StartIndex + 4).Y = outVerts(StartIndex + 4).Y + Y
    
    outVerts(StartIndex + 3).X = outVerts(StartIndex + 1).X
    outVerts(StartIndex + 3).Y = outVerts(StartIndex + 1).Y
    outVerts(StartIndex + 5).X = outVerts(StartIndex + 2).X
    outVerts(StartIndex + 5).Y = outVerts(StartIndex + 2).Y
    
End Sub

Public Function GUIImageWidth(ByVal ID As Integer) As Integer
    GUIImageWidth = Image(ID).W
End Function

Public Function GUIImageHeight(ByVal ID As Integer) As Integer
    GUIImageHeight = Image(ID).H
End Function

Public Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

'Olvidé guardar en el binario la cantidad de elementos, asi que los obtengo asi :^)

Private Function NumbersOfElements(ByVal FileSize As Long) As Long
    Dim g As tGUIImages
    
    NumbersOfElements = FileSize \ LenB(g)
    
End Function

Private Sub FillVerts()
Dim I As Long
Dim VertsCount As Long
Dim Width As Integer, Height As Integer

    Width = GUITex.Width
    Height = GUITex.Height
    
    VertsCount = 0
    For I = 0 To UBound(Image)
        
        Verts(VertsCount).X = 0
        Verts(VertsCount).Y = 0
        Verts(VertsCount).tu = Image(I).X / Width
        Verts(VertsCount).tv = Image(I).Y / Height
        'Verts(VertsCount).Rhw = 1
        Verts(VertsCount).Color = -1
        
        Verts(VertsCount + 1).X = Image(I).W
        Verts(VertsCount + 1).Y = 0
        Verts(VertsCount + 1).tu = (Image(I).X + Image(I).W) / Width
        Verts(VertsCount + 1).tv = Image(I).Y / Height
        'Verts(VertsCount + 1).Rhw = 1
        Verts(VertsCount + 1).Color = -1
        
        Verts(VertsCount + 2).X = 0
        Verts(VertsCount + 2).Y = Image(I).H
        Verts(VertsCount + 2).tu = Image(I).X / Width
        Verts(VertsCount + 2).tv = (Image(I).Y + Image(I).H) / Height
        'Verts(VertsCount + 2).Rhw = 1
        Verts(VertsCount + 2).Color = -1
        
        Verts(VertsCount + 4).X = Image(I).W
        Verts(VertsCount + 4).Y = Image(I).H
        Verts(VertsCount + 4).tu = (Image(I).X + Image(I).W) / Width
        Verts(VertsCount + 4).tv = (Image(I).Y + Image(I).H) / Height
        'Verts(VertsCount + 4).Rhw = 1
        Verts(VertsCount + 4).Color = -1
        
        Verts(VertsCount + 3) = Verts(VertsCount + 1)
        Verts(VertsCount + 5) = Verts(VertsCount + 2)
        
        VertsCount = VertsCount + 6
    Next

End Sub
