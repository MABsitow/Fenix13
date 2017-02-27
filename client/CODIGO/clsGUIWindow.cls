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

Private Rectangle As Rect

Private HasTitle As Boolean
Private CloseButton As Boolean
Private StrTitle As String

Private Objects() As clsGUIObject
Private ObjectsCount As Long

Private ObjectFocused As Integer

Public Sub Initialize(ByVal X As Integer, ByVal Y As Integer, _
                            ByVal Width As Integer, ByVal Height As Integer, _
                            Optional ByVal Title As String = "", Optional ByVal CloseBtn As Boolean = True)
    
    With Rectangle
        .x1 = X
        .x2 = Width
        .y1 = Y
        .y2 = Height
    End With
    
    'Resize the window to a multiple of 32 'CHECK THIS
    Call NormalizeSize
    
    If StrComp(Title, vbNullString) > 0 Then
        HasTitle = True
        StrTitle = Title
    End If
    
    CloseButton = CloseBtn

End Sub

Public Sub OnDraw(Batch As clsBGFXSpriteBatch, ByVal AtlasW As Long, ByVal AtlasH As Long)
    'todo: Would be great to build an image in memory and draw it at once insted of draw it by pieces
    'todo: GUI should not use AmbientColor
    
    Dim i As Long
    Dim j As Long
    
    Dim Obj As clsGUIObject
    
    Dim startX As Long
    Dim startY As Long
    
    Dim compW As Integer
    
    startX = Rectangle.x1 'should check for w & h (not render, or cut the images)
    startY = Rectangle.y1
    
    For i = 1 To ObjectsCount
        Set Obj = Objects(i)
        
        Select Case Obj.OType
        
            Case eGuiObjectType.Button
                
                With Image(eGuiElement.LeftButton)
                
                    Call Batch.Draw(startX + Obj.X, startY + Obj.Y, _
                                    .W, .H, _
                                    AmbientColor, _
                                    .X / AtlasW, .Y / AtlasH, _
                                    (.X + .W) / AtlasW, (.Y + .H) / AtlasH)
                        
                End With
                
                startX = startX + GUIImageWidth(eGuiElement.LeftButton)
                
                compW = GUIImageWidth(eGuiElement.MidButton)
                
                With Image(eGuiElement.MidButton)

                    For j = 0 To (Obj.W / compW) - 3
                        
                        Call Batch.Draw(startX + Obj.X, startY + Obj.Y, _
                                        .W, .H, _
                                        AmbientColor, _
                                        .X / AtlasW, .Y / AtlasH, _
                                        (.X + .W) / AtlasW, (.Y + .H) / AtlasH)
                        
                        startX = startX + compW
                    Next
                    
                End With
                
                With Image(eGuiElement.RightButton)
                
                    Call Batch.Draw(startX + Obj.X, startY + Obj.Y, _
                                    .W, .H, _
                                    AmbientColor, _
                                    .X / AtlasW, .Y / AtlasH, _
                                    (.X + .W) / AtlasW, (.Y + .H) / AtlasH)
                
                End With
        End Select
    Next
    
End Sub

Public Sub AddObject(ByVal Name As String, ByVal X As Integer, ByVal Y As Integer, _
                    ByVal W As Integer, ByVal H As Integer, ByVal OType As eGuiObjectType, events() As Long, _
                    ByVal Text As String)

Dim GuiObject As New clsGUIObject
Dim Rect As Rect
    
    With Rect
        .x1 = X
        .x2 = W
        .y1 = Y
        .y2 = H
    End With
    
    Call GuiObject.CreateObject(Name, Rect.x1, Rect.y1, Rect.x2, Rect.y2, OType, events, Text)
    
    ObjectsCount = ObjectsCount + 1
    ReDim Preserve Objects(1 To ObjectsCount) As clsGUIObject
    Set Objects(ObjectsCount) = GuiObject

End Sub

Friend Function getObject(ByVal Index As Integer) As clsGUIObject
    Set getObject = Objects(Index)
End Function

Friend Function getObjectCount() As Integer
    getObjectCount = ObjectsCount
End Function

Friend Function Collide(ByVal X As Integer, ByVal Y As Integer) As Integer
Dim i As Long

Dim comp As clsGUIObject

For i = 1 To ObjectsCount
    Set comp = Objects(i)

    If X > (comp.X + Rectangle.x1) And X < (comp.W + Rectangle.x1 + comp.X) And _
        Y > (comp.Y + Rectangle.y1) And Y < (comp.H + Rectangle.y1 + comp.Y) Then
            
            Collide = i
            Exit Function
    End If
Next

Collide = -1
End Function

Function Engine_Collision_Rect(ByVal x1 As Integer, ByVal y1 As Integer, ByVal Width1 As Integer, ByVal Height1 As Integer, ByVal x2 As Integer, ByVal y2 As Integer, ByVal Width2 As Integer, ByVal Height2 As Integer) As Boolean
'*****************************************************************
'Check for collision between two rectangles
'More info: http://www.vbgore.com/GameClient.TileEngine.Engine_Collision_Rect
'*****************************************************************

    If x1 + Width1 >= x2 Then
        If x1 <= x2 + Width2 Then
            If y1 + Height1 >= y2 Then
                If y1 <= y2 + Height2 Then
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
    
    Do While Rectangle.y2 Mod 32 <> 0
        Rectangle.y2 = Rectangle.y2 + 1
    Loop
    
End Sub
