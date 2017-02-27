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

'atlas
Private GUITex As TYPE_VIDEO_IMAGE

'files
Private Type tGUIImages
        ID As Long
        X As Long
        Y As Long
        W As Long
        H As Long
End Type

Public Image() As tGUIImages

Private frm As New clsGUIWindow 'test window

Public Sub LoadGUIImage()
    Dim handle As Integer
    Dim elements As Integer
    
    handle = FreeFile
    
    Open IniPath & "gui.fnx" For Binary As handle
        Seek handle, 1
        
        elements = NumbersOfElements(LOF(handle))
        
        ReDim Image(0 To elements - 1) As tGUIImages
        
        Get handle, , Image
        
    Close handle
    
    GUITex = Video.CreateImageFromFilename(DirGraficos & "Gui.png")
    
    Call frm.Initialize(50, 50, 320, 320)
    
    Dim events() As Long
    
    Call frm.AddObject("Lbl1", 25, 25, 0, 0, eGuiObjectType.Label, events, "Version: FXIII")
    Call frm.AddObject("Btn1", 50, 50, 96, 53, eGuiObjectType.Button, events, "Hola")
    Call frm.AddObject("Lbl2", 50, 300, 0, 0, eGuiObjectType.Label, events, "LL")
    
End Sub

Public Sub DrawAllWindows(Batch As clsBGFXSpriteBatch)
    
    Call Batch.SetTexture(GUITex.mHandle)
    
    Call frm.OnDraw(Batch, GUITex.mX, GUITex.mY)
    
End Sub

Public Sub GuiMouseMove(ByVal X As Integer, ByVal Y As Integer)
'te extraño for each

Dim i As Long

'For i = 0 To windowsCount
    
    i = frm.Collide(X, Y)
    If i <> -1 Then
    
        'Call Text_Colorize(frm.getObject(I).TSV, Len(frm.getObject(I).Text), D3DColorXRGB(255, 0, 0))
        
    End If
    
    
End Sub

Public Function GUIImageWidth(ByVal ID As Integer) As Integer
    GUIImageWidth = Image(ID).W
End Function

Public Function GUIImageHeight(ByVal ID As Integer) As Integer
    GUIImageHeight = Image(ID).H
End Function


'Olvidé guardar en el binario la cantidad de elementos, asi que los obtengo asi :^)

Private Function NumbersOfElements(ByVal FileSize As Long) As Long
    Dim G As tGUIImages
    
    NumbersOfElements = FileSize \ LenB(G)
    
End Function
