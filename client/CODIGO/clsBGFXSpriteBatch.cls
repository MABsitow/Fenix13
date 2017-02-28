VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBGFXSpriteBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
' This file is part of BGFXBridgeVB6, licensed under the APACHE License.
'
' Copyright (c) 2017 Agustin L. Alvarez <wolftein1@gmail.com>
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
' http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'

Option Explicit

'
' API
'
Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, Source As Any, ByVal Length As Long)

'
' 2D vertex structure
'
Private Type TYPE_VERTEX
            
    X       As Single
    Y       As Single
    Z       As Single
    
    Colour  As Long
    
    TX      As Single
    TY      As Single

End Type

'
' Current view of the batch
'
Private m_View    As Byte

'
' Current texture being use for the batch
'
Private m_Texture As Long

'
' Current program being use for the batch
'
Private m_Program As Long

'
' Current layout being use for the batch
'
Private m_Layout  As TYPE_VIDEO_LAYOUT

'
' IBuffer handle of the batch (statically compiled)
'
Private m_Indices As Long

'
' VBuffer handle of the batch (dynamically compiled)
'
Private m_Vertices As TYPE_VIDEO_TRANSIENT_VERTEX_BUFFER

'
' Current position of the batch (in vertices coordinates)
'
Private m_Position As Long

'
' Current data of the batch (VB6 doesn't support casting, so no direct mapping unless we hack it)
'
Private m_Data() As TYPE_VERTEX

'
' Current states
'
Private m_State As TYPE_VIDEO_STATE


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Begin()
    
    '
    ' Nothing to see in here
    '
    
End Sub


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Draw(ByVal X As Long, _
                ByVal Y As Long, _
                ByVal Z As Long, _
                ByVal Width As Long, _
                ByVal Height As Long, _
                ByRef Color() As Long, _
                Optional ByVal Tx1 As Single = 0, _
                Optional ByVal Ty1 As Single = 0, _
                Optional ByVal Tx2 As Single = 1, _
                Optional ByVal Ty2 As Single = 1)

    m_Data(m_Position).X = X
    m_Data(m_Position).Y = Y
    m_Data(m_Position).Z = Z
    m_Data(m_Position).Colour = Color(0)
    m_Data(m_Position).TX = Tx1
    m_Data(m_Position).TY = Ty1
    
    m_Position = m_Position + 1
    
    m_Data(m_Position).X = X
    m_Data(m_Position).Y = Y + Height
    m_Data(m_Position).Z = Z
    m_Data(m_Position).Colour = Color(1)
    m_Data(m_Position).TX = Tx1
    m_Data(m_Position).TY = Ty2
    
    m_Position = m_Position + 1
    
    m_Data(m_Position).X = X + Width
    m_Data(m_Position).Y = Y + Height
    m_Data(m_Position).Z = Z
    m_Data(m_Position).Colour = Color(2)
    m_Data(m_Position).TX = Tx2
    m_Data(m_Position).TY = Ty2
    
    m_Position = m_Position + 1
    
    m_Data(m_Position).X = X + Width
    m_Data(m_Position).Y = Y
    m_Data(m_Position).Z = Z
    m_Data(m_Position).Colour = Color(3)
    m_Data(m_Position).TX = Tx2
    m_Data(m_Position).TY = Ty1
    
    m_Position = m_Position + 1
    
    If (m_Position = UBound(m_Data) + 1) Then
        Call Flush
    End If
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Finish()
    
    '
    ' Flush all primitives before finishing batching
    '
    Call Flush
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Flush()
                
    '
    ' Only proceed if we have something to draw
    '
    If m_Position > 0 Then
        '
        ' Update
        '
        m_Vertices = Video.CreateTransientVertexBuffer(m_Position, m_Layout)
        
        Call RtlMoveMemory(ByVal m_Vertices.mPointer, m_Data(0), 24 * m_Position)
        
        '
        ' Apply states
        '
        Call Video.SetState(m_State)
        
        ' Apply texture state
        '
        Call Video.SetTexture(0, 0, m_Texture)
                
        '
        ' Apply index buffer state
        '
        Call Video.SetIndexBuffer(m_Indices, 0, (m_Position / 4) * 6)
        
        '
        ' Apply vertex buffer state
        '
        Call Video.SetTransientVertexBuffer(m_Vertices, 0, m_Position)
    
        '
        ' Submit all primitives to the target view
        '
        Call Video.Submit(m_View, m_Program)
    
        '
        ' Reset the position of the batch
        '
        m_Position = 0
        
    End If

End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub Initialise(ByVal Capacity As Long)

    '
    ' Create the vertex's format
    '
    Call Video.VertexDeclBegin(m_Layout)
    Call Video.VertexDeclAdd(m_Layout, ATTRIBUTE_POSITION, 3, ATTRIBUTE_TYPE_FLOAT)
    Call Video.VertexDeclAdd(m_Layout, ATTRIBUTE_COLOR0, 4, ATTRIBUTE_TYPE_UINT8, True)
    Call Video.VertexDeclAdd(m_Layout, ATTRIBUTE_TEXCOORD0, 2, ATTRIBUTE_TYPE_FLOAT)
    Call Video.VertexDeclEnd(m_Layout)
    
    '
    ' Pre create the data array
    '
    ReDim m_Data(0 To Capacity * 4) As TYPE_VERTEX
    
    '
    ' Pre calculate all indices for the batch and create the buffer
    '
    ReDim wData(0 To Capacity * 6 - 1) As Integer
    
    Dim i As Integer, j As Integer
   
    For i = 0 To UBound(wData) Step 6
        wData(i) = j
        wData(i + 1) = j + 1
        wData(i + 2) = j + 2
        wData(i + 3) = j + 2
        wData(i + 4) = j + 3
        wData(i + 5) = j
       
        j = j + 4
    Next

    m_Indices = Video.CreateIndexBufferFromMemory(Video.Copy16I(wData))
    
    
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetProgram(ByVal Program As Long)

    If (m_Program <> Program) Then
        '
        ' Flush the current batch before changing program state
        '
        Call Flush
        
        m_Program = Program
    End If
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetTexture(ByVal Texture As Long)

    If (m_Texture <> Texture) Then
        '
        ' Flush the current batch before changing texture state
        '
        Call Flush
        
        m_Texture = Texture
    End If
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetView(ByVal View As Byte)

    If (m_View <> View) Then
        '
        ' Flush the current batch before changing view state
        '
        Call Flush
        
        m_View = View
    End If
        
End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub SetState(State As TYPE_VIDEO_STATE)
    
    If (m_State.mMask(0) <> State.mMask(0) Or _
        m_State.mMask(1) <> State.mMask(1)) Then
        
        '
        ' Flush the current batch before changing states
        '
        Call Flush
        
        m_State = State
    End If
    
End Sub

