VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGUIWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Based on: http://www.gs-zone.org/temas/gui.84864/

Option Explicit

Private Rectangle As D3DRect

Private HasTitle As Boolean
Private CloseButton As Boolean
Private StrTitle As String

Private Verts() As TLVERTEX
Private VertsCount As Long

Private ElementsCount As Long

Private Objects() As clsGUIObject
Private ObjectsCount As Long

Private ObjectFocused As Integer

Public Sub Initialize(ByVal X As Integer, ByVal Y As Integer, _
                            ByVal Width As Integer, ByVal Height As Integer, _
                            Optional ByVal Title As String = "", Optional ByVal CloseBtn As Boolean = True)
    
    With Rectangle
        .x1 = X
        .x2 = Width
        .Y1 = Y
        .Y2 = Height
    End With
    
    'Resize the window to a multiple of 32 'CHECK THIS
    Call NormalizeSize
    
    If StrComp(Title, vbNullString) > 0 Then
        HasTitle = True
        StrTitle = Title
    End If
    
    CloseButton = CloseBtn
    
    ReDim Verts(0 To 1000) As TLVERTEX
    
    Call BuildWindowVerts
End Sub

Public Sub OnDraw(Batch As clsBatch)
    
    Dim i As Long
    
    For i = 1 To ObjectsCount
        
    Next
    
End Sub

Public Sub AddObject(ByVal Name As String, ByVal X As Integer, ByVal Y As Integer, _
                    ByVal W As Integer, ByVal H As Integer, ByVal OType As eGuiObjectType, events() As Long, _
                    ByVal Text As String, Optional BuildObjects As Boolean = False)

'param BuilObjects: Solo debe ser true cuando agreguen el último objecto a la ventana

Dim GuiObject As New clsGUIObject
Dim Rect As D3DRect
    
    With Rect
        .x1 = X
        .x2 = W
        .Y1 = Y
        .Y2 = H
    End With
    
    Call GuiObject.CreateObject(Name, Rect, OType, events, Text)
    
    ObjectsCount = ObjectsCount + 1
    ReDim Preserve Objects(1 To ObjectsCount) As clsGUIObject
    Set Objects(ObjectsCount) = GuiObject
    
    
    If BuildObjects Then Call AddObjectsVerts
End Sub

Private Sub AddObjectsVerts()
Dim i As Long, j As Long
Dim startX As Integer, startY As Integer
Dim compW As Integer, compH As Integer

'agrego mas vertices para poder trabajar
ReDim Preserve Verts(0 To VertsCount + 300) As TLVERTEX

Dim comp As clsGUIObject

For i = 1 To ObjectsCount
    
    Set comp = Objects(i)
    
    'Nos posicionamos dentro de la ventana
    startX = Rectangle.x1
    startY = Rectangle.Y1

    With comp
        Select Case comp.OType
            Case eGuiObjectType.Label
                If i = 0 Then
                    VertsCount = VertsCount + 1
                End If
                
                startX = Rectangle.x1 + comp.ObjectRectangle.x1
                startY = Rectangle.Y1 + comp.ObjectRectangle.Y1
                
                Call Text_Add(comp.Text, startX, startY, -1)
                
            Case eGuiObjectType.Button
            
                If i = 0 Then ' evito pisar el ultimo vertice cargado
                    VertsCount = VertsCount + 1
                End If
                
                startX = startX + comp.ObjectRectangle.x1
                startY = startY + comp.ObjectRectangle.Y1
                
                Dim txtX As Integer
                Dim txtY As Integer
                
                txtX = startX + (comp.ObjectRectangle.x2 \ 2) - (Text_GetWidth(cfonts(1), comp.Text) \ 2)
                txtY = startY + (comp.ObjectRectangle.Y2 \ 2) - (cfonts(1).RowPitch \ 2)
                
                comp.TSV = Text_Add(comp.Text, txtX, txtY, -1)
                
                Call VertsFromElement(eGuiElement.LeftButton, Verts, VertsCount)
                Call Translate(Verts, VertsCount, startX, startY)
                VertsCount = VertsCount + 6
                ElementsCount = ElementsCount + 1
                
                startX = startX + GUIImageWidth(eGuiElement.LeftButton)
                
                compW = GUIImageWidth(eGuiElement.MidButton)
                For j = 0 To (comp.ObjectRectangle.x2 / compW) - 3
                    Call VertsFromElement(eGuiElement.MidButton, Verts, VertsCount)
                    Call Translate(Verts, VertsCount, startX, startY)
                    VertsCount = VertsCount + 6
                    ElementsCount = ElementsCount + 1
                    startX = startX + GUIImageWidth(eGuiElement.MidButton)
                Next
                
                Call VertsFromElement(eGuiElement.RightButton, Verts, VertsCount)
                Call Translate(Verts, VertsCount, startX, startY)
                VertsCount = VertsCount + 6
                ElementsCount = ElementsCount + 1
                
        End Select
    End With
    
Next

'volvemos a ajustar
ReDim Preserve Verts(0 To VertsCount) As TLVERTEX

End Sub
Private Sub BuildWindowVerts()
'Prepare the verts
'Check?
    Dim startX As Integer, startY As Integer
    Dim elementW As Integer, elementH As Integer
    
    startX = Rectangle.x1
    startY = Rectangle.Y1
    
    'Arriba izquierda
    Call VertsFromElement(eGuiElement.TopLeft, Verts)
    Call Translate(Verts, VertsCount, startX, startY)
    
    startX = startX + GUIImageWidth(eGuiElement.TopLeft)
    
    VertsCount = VertsCount + 6
    ElementsCount = ElementsCount + 1
    
    Dim i As Long
    Dim j As Long
    
    elementW = GUIImageWidth(eGuiElement.TopMid)
    'Arriba al medio
    For i = 0 To (Rectangle.x2 / elementW) - 3
        Call VertsFromElement(eGuiElement.TopMid, Verts, VertsCount)
        Call Translate(Verts, VertsCount, startX, startY)
        VertsCount = VertsCount + 6
        ElementsCount = ElementsCount + 1
        startX = startX + GUIImageWidth(eGuiElement.TopMid)
    Next
    
    'Arriba derecha
    Call VertsFromElement(eGuiElement.TopRight, Verts, VertsCount)
    Call Translate(Verts, VertsCount, startX, startY)
    VertsCount = VertsCount + 6
    ElementsCount = ElementsCount + 1
    
    'reacomodo
    startX = Rectangle.x1
    startY = startY + GUIImageHeight(eGuiElement.TopLeft)
    
    elementH = GUIImageHeight(eGuiElement.LeftMid)
    'izquierda medio
    For i = 0 To (Rectangle.Y2 / elementH) - 3
        Call VertsFromElement(eGuiElement.LeftMid, Verts, VertsCount)
        Call Translate(Verts, VertsCount, startX, startY)
        VertsCount = VertsCount + 6
        ElementsCount = ElementsCount + 1
        startY = startY + GUIImageHeight(eGuiElement.LeftMid)
    Next
    
    'abajo izquierda
    Call VertsFromElement(eGuiElement.BotLeft, Verts, VertsCount)
    Call Translate(Verts, VertsCount, startX, startY)
    VertsCount = VertsCount + 6
    ElementsCount = ElementsCount + 1
    startX = startX + GUIImageWidth(eGuiElement.BotLeft)
    
    elementW = GUIImageWidth(eGuiElement.BotMid)
    'abajo al medio
    For i = 0 To (Rectangle.x2 / elementW) - 3
        Call VertsFromElement(eGuiElement.BotMid, Verts, VertsCount)
        Call Translate(Verts, VertsCount, startX, startY)
        VertsCount = VertsCount + 6
        ElementsCount = ElementsCount + 1
        startX = startX + GUIImageWidth(eGuiElement.BotMid)
    Next
    
    Call VertsFromElement(eGuiElement.BotRight, Verts, VertsCount)
    Call Translate(Verts, VertsCount, startX, startY)
    VertsCount = VertsCount + 6
    ElementsCount = ElementsCount + 1
    startY = Rectangle.Y1 + GUIImageHeight(eGuiElement.BotRight)
    
    elementH = GUIImageHeight(eGuiElement.RightMid)
    'derecho al medio
    For i = 0 To (Rectangle.Y2 / elementH) - 3
        Call VertsFromElement(eGuiElement.RightMid, Verts, VertsCount)
        Call Translate(Verts, VertsCount, startX, startY)
        VertsCount = VertsCount + 6
        startY = startY + GUIImageHeight(eGuiElement.RightMid)
        ElementsCount = ElementsCount + 1
    Next
    
    'reacomodo
    startX = Rectangle.x1 + GUIImageWidth(eGuiElement.LeftMid)
    startY = Rectangle.Y1 + GUIImageHeight(eGuiElement.TopMid)
    
    elementW = GUIImageWidth(eGuiElement.BG)
    elementH = GUIImageHeight(eGuiElement.BG)
    'fondo
    For i = 0 To (Rectangle.x2 / elementW) - 3
        For j = 0 To (Rectangle.Y2 / elementH) - 3
            
            Call VertsFromElement(eGuiElement.BG, Verts, VertsCount)
            Call Translate(Verts, VertsCount, startX, startY)
            VertsCount = VertsCount + 6
            startY = startY + GUIImageHeight(eGuiElement.BG)
            ElementsCount = ElementsCount + 1
        Next
        startY = Rectangle.Y1 + GUIImageHeight(eGuiElement.TopMid) 'reset
        startX = startX + GUIImageWidth(eGuiElement.BG)
    Next
    
    'redim para salvar memoria - probablemente nunca vaya a tener una ventana sin controles _
    por lo cual esto se podria evitar y hacerlo cuando se carguen todos los componentes _
    PEEEEEEERO algun loco puede llegar a crear la ventana sin componentes y terminaria _
    ocupando mas memoria de la que debe si esto no estuviera
    
    ReDim Preserve Verts(0 To VertsCount) As TLVERTEX
    
End Sub

Friend Function getObject(ByVal Index As Integer) As clsGUIObject
    Set getObject = Objects(Index)
End Function

Friend Function getObjectCount() As Integer
    getObjectCount = ObjectsCount
End Function

Friend Function getElementsCount() As Integer
    getElementsCount = ElementsCount
End Function

Friend Function getVertsCount() As Integer
    getVertsCount = VertsCount
End Function

Friend Sub getWindowVerts(ByRef outVerts() As TLVERTEX)
    CopyMemory outVerts(0), Verts(0), FVF_SIZE * VertsCount
End Sub

Friend Function Collide(ByVal X As Integer, ByVal Y As Integer) As Integer
Dim i As Long

Dim comp As clsGUIObject
Dim Rect As D3DRect

For i = 1 To ObjectsCount
    Set comp = Objects(i)
    Rect = comp.ObjectRectangle
    
    With Rect
        If X > (.x1 + Rectangle.x1) And X < (.x2 + Rectangle.x1 + .x1) And _
            Y > (.Y1 + Rectangle.Y1) And Y < (.Y2 + Rectangle.Y1 + .Y1) Then
                
                Collide = i
                Exit Function
        End If
    
    End With
Next

Collide = -1
End Function

Function Engine_Collision_Rect(ByVal x1 As Integer, ByVal Y1 As Integer, ByVal Width1 As Integer, ByVal Height1 As Integer, ByVal x2 As Integer, ByVal Y2 As Integer, ByVal Width2 As Integer, ByVal Height2 As Integer) As Boolean
'*****************************************************************
'Check for collision between two rectangles
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_Rect
'*****************************************************************

    If x1 + Width1 >= x2 Then
        If x1 <= x2 + Width2 Then
            If Y1 + Height1 >= Y2 Then
                If Y1 <= Y2 + Height2 Then
                    Engine_Collision_Rect = True
                End If
            End If
        End If
    End If

End Function

Private Sub NormalizeSize()

    Do While Rectangle.x2 Mod 32 <> 0
        Rectangle.x2 = Rectangle.x2 + 1
    Loop
    
    Do While Rectangle.Y2 Mod 32 <> 0
        Rectangle.Y2 = Rectangle.Y2 + 1
    Loop
    
End Sub
