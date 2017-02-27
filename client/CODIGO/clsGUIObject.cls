VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGUIObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum eGuiEvents
    None 'estado normal
    MouseDown 'es clickeado
    MouseMove 'se pasa el mouse por encima
    KeyDown
End Enum

Private Rectangle       As Rect
Private OEvents()       As Long 'eventos

Public Name             As String
Public OType            As eGuiObjectType 'tipo de objeto

Public Text             As String

Private Const MIN_BUTTON_WIDTH As Integer = 96

Public Sub CreateObject(ByVal ObjectName As String, X As Long, Y As Long, W As Long, H As Long, ByVal ObjectType As eGuiObjectType, events() As Long, ByVal ObjectText As String)
    Name = ObjectName
    
    With Rectangle
        .x1 = X
        .x2 = W
        .y1 = Y
        .y2 = H
    End With
    
    OType = ObjectType
    
    If OType = eGuiObjectType.Button Then Call Normalize
    
    Text = ObjectText
    
    'Si events es nulo, no hacemos un redim
    If Peek(ArrPtr(events)) <> 0 Then
        ReDim OEvents(0 To UBound(events)) As Long
    
        OEvents = events
    End If
End Sub

Private Function Peek(ByVal lPtr As Long) As Long
    Call CopyMemory(Peek, ByVal lPtr, 4)
End Function

Private Sub Normalize()
    
    'Buttons are width normalized (height is a fixed size)
    
    Do While (Rectangle.x2 Mod 32 <> 0) And (Rectangle.x2 < MIN_BUTTON_WIDTH)
        Rectangle.x2 = Rectangle.x2 + 1
    Loop
    
End Sub
Public Property Get X() As Long
    X = Rectangle.x1
End Property

Public Property Get Y() As Long
    Y = Rectangle.y1
End Property

Public Property Get W() As Long
    W = Rectangle.x2
End Property

Public Property Get H() As Long
    H = Rectangle.y2
End Property

Public Property Get ObjectEvents() As Long()
    ObjectEvents = OEvents
End Property
