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

Private Rectangle       As D3DRect
Private OEvents()       As Long 'eventos

Public Name             As String
Public OType            As eGuiObjectType 'tipo de objeto

Public Text             As String
Public TSV              As Long 'Text Start Vert

Private Const MIN_BUTTON_WIDTH As Integer = 96

Public Sub CreateObject(ByVal ObjectName As String, Rect As D3DRect, ByVal ObjectType As eGuiObjectType, events() As Long, ByVal ObjectText As String)
    Name = ObjectName
    Rectangle = Rect
    OType = ObjectType
    
    If OType = eGuiObjectType.Button Then Call Normalize
    
    Text = ObjectText
    
    'Si events es nulo, no hacemos un redim
    If Peek(ArrPtr(events)) <> 0 Then
        ReDim OEvents(0 To UBound(events)) As Long
    
        OEvents = events
    End If
End Sub

Private Sub Normalize()
    
    'Buttons are width normalized (height is a fixed size)
    
    Do While (Rectangle.x2 Mod 32 <> 0) And (Rectangle.x2 < MIN_BUTTON_WIDTH)
        Rectangle.x2 = Rectangle.x2 + 1
    Loop
    
End Sub
Public Property Get ObjectRectangle() As D3DRect
    ObjectRectangle = Rectangle
End Property

Public Property Get ObjectEvents() As Long()
    ObjectEvents = OEvents
End Property
