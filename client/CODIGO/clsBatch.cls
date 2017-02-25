VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBatch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
 
Private m_VBuffer  As Direct3DVertexBuffer8
Private m_IBuffer  As Direct3DIndexBuffer8
Private m_Texture  As Direct3DTexture8
 
'
' Represent the dynamic data of the batch
'
Private m_Data()   As Single
Private m_Position As Long
 
Public Sub Begin()
    '
    ' Re-initialise the states of the batch
    '
    Call DirectDevice.SetTexture(0, m_Texture)
    Call DirectDevice.SetStreamSource(0, m_VBuffer, 24)
    Call DirectDevice.SetIndices(m_IBuffer, 0)
End Sub

Public Sub SetTexture(ByRef Texture As Direct3DTexture8)

    If ObjPtr(m_Texture) <> ObjPtr(Texture) Then 'aca cague jaja te cabe vb6
        '
        ' Flush the batch since we're using a diferent texture
        '
        Call Flush
       
        Set m_Texture = Texture
       
        '
        ' Update the new texture of the batch
        '
        Call DirectDevice.SetTexture(0, Texture)
    End If
End Sub
 
Public Sub Draw(ByVal X As Long, _
                ByVal Y As Long, _
                ByVal Width As Long, _
                ByVal Height As Long, _
                ByVal Color As Long, _
                Optional ByVal Tx1 As Single = 0, _
                Optional ByVal Ty1 As Single = 0, _
                Optional ByVal Tx2 As Single = 1, _
                Optional ByVal Ty2 As Single = 1)

    m_Data(m_Position) = X
    m_Position = m_Position + 1
    m_Data(m_Position) = Y
    m_Position = m_Position + 1
    m_Data(m_Position) = 0 ' Z
    m_Position = m_Position + 1
    m_Data(m_Position) = Color
    m_Position = m_Position + 1
    m_Data(m_Position) = Tx1
    m_Position = m_Position + 1
    m_Data(m_Position) = Ty1
    m_Position = m_Position + 1
   
    m_Data(m_Position) = X + Width
    m_Position = m_Position + 1
    m_Data(m_Position) = Y
    m_Position = m_Position + 1
    m_Data(m_Position) = 0 ' Z
    m_Position = m_Position + 1
    m_Data(m_Position) = Color
    m_Position = m_Position + 1
    m_Data(m_Position) = Tx2
    m_Position = m_Position + 1
    m_Data(m_Position) = Ty1
    m_Position = m_Position + 1

    m_Data(m_Position) = X
    m_Position = m_Position + 1
    m_Data(m_Position) = Y + Height
    m_Position = m_Position + 1
    m_Data(m_Position) = 0 ' Z
    m_Position = m_Position + 1
    m_Data(m_Position) = Color
    m_Position = m_Position + 1
    m_Data(m_Position) = Tx1
    m_Position = m_Position + 1
    m_Data(m_Position) = Ty2
    m_Position = m_Position + 1
           
    m_Data(m_Position) = X + Width
    m_Position = m_Position + 1
    m_Data(m_Position) = Y + Height
    m_Position = m_Position + 1
    m_Data(m_Position) = 0 ' Z
    m_Position = m_Position + 1
    m_Data(m_Position) = Color
    m_Position = m_Position + 1
    m_Data(m_Position) = Tx2
    m_Position = m_Position + 1
    m_Data(m_Position) = Ty2
    m_Position = m_Position + 1
   
    If (m_Position = UBound(m_Data)) Then
        Call Flush
    End If
End Sub
 
Public Sub Finish()
    '
    ' Flush the batch since we're finished
    '
    Call Flush
End Sub
 
Public Sub Flush()

    '
    ' Only draw if there is any data in the buffer
    '
    If m_Position > 0 Then
        '
        ' Set the new data of the vertex buffer
        '
        Call D3DVertexBuffer8SetData(m_VBuffer, 0, m_Position * 4, 0, m_Data(0))
       
        Call DirectDevice.DrawIndexedPrimitive(D3DPT_TRIANGLELIST, 0, (m_Position / 24) * 6, 0, (m_Position / 24) * 2)
    
        '
        ' We start all over again
        '
        m_Position = 0
    End If
End Sub
 
Public Sub Initialise(ByVal Capacity As Long)
    ReDim m_Data(24 * Capacity) As Single
   
    '
    '  Create the vertice buffer
    '
    Set m_VBuffer = DirectDevice.CreateVertexBuffer(24 * Capacity, D3DUSAGE_DYNAMIC, D3DFVF_XYZ Or D3DFVF_DIFFUSE Or D3DFVF_TEX1, D3DPOOL_DEFAULT)
  
    '
    '  Create the indice buffer, fill it with pre-baked indices
    '
    Set m_IBuffer = DirectDevice.CreateIndexBuffer(12 * Capacity, D3DUSAGE_WRITEONLY, D3DFMT_INDEX16, D3DPOOL_DEFAULT)
   
    Dim lpIndices() As Integer
    
    ReDim lpIndices(0 To Capacity * 6) As Integer
    
    Dim i As Integer, j As Integer
   
    For i = 0 To UBound(lpIndices) - 1 Step 6
        lpIndices(i) = j
        lpIndices(i + 1) = j + 1
        lpIndices(i + 2) = j + 2
        lpIndices(i + 3) = j + 2
        lpIndices(i + 4) = j + 3
        lpIndices(i + 5) = j
       
        j = j + 4
    Next

    Call D3DIndexBuffer8SetData(m_IBuffer, 0, UBound(lpIndices), 0, lpIndices(0))
End Sub
