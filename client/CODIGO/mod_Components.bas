Attribute VB_Name = "mod_Components"
Option Explicit

Private Const MAX_CONSOLE_LINES As Byte = 100

Public UserWriting As Boolean

Public Enum eComponentEvent
        None = 0
        MouseMove = 1
        MouseDown = 2
        KeyUp = 3
        KeyPress = 4
        MouseScrollUp = 5
        MouseScrollDown = 6
End Enum

Public Enum eComponentType
        Label = 0
        TextBox = 1
        Shape = 2
        TextArea = 3
End Enum

Private Type TYPE_CONSOLE_LINE
        Text As String
        Color(3) As Long
End Type

Private Type tComponent 'todo: rehacer?
        X           As Integer
        Y           As Integer
        W           As Integer
        H           As Integer
        
        Component   As eComponentType
        
        IsFocusable As Boolean
        ShowOnFocus As Boolean 'Only showed when its focused
        Color(3)    As Long
        
        Text        As String
        TextBuffer  As String 'Buffer
        
        ForeColor(3) As Long
        
        EventsPtr   As Long
        HasEvents   As Boolean
        
        'TextArea
        Lines()     As TYPE_CONSOLE_LINE
        LastLine    As Byte
        
        'first and last line to render in console
        FirstRender As Byte
        LastRender  As Byte
End Type

Private BackgroundImage As TYPE_VIDEO_IMAGE

Private Focused         As Integer
Private LastComponent   As Integer

Public Components()     As tComponent

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


Public Sub InitComponentsImage()
    
    BackgroundImage = Video.CreateImageFromFilename(DirGraficos & "4.png")
    
End Sub

Public Function AddTextArea(ByVal X As Integer, ByVal Y As Integer, _
                            ByVal W As Integer, ByVal H As Integer, _
                            Color() As Long) As Integer
    
    LastComponent = LastComponent + 1
    
    ReDim Preserve Components(1 To LastComponent) As tComponent
    
    With Components(LastComponent)
    
        .X = X: .W = W
        .Y = Y: .H = H
        
        .Component = eComponentType.TextArea
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
    End With
    
    AddTextArea = LastComponent
    
End Function

Public Function AddLabel(Text As String, ByVal X As Integer, ByVal Y As Integer, Color() As Long) As Integer

    LastComponent = LastComponent + 1
    
    ReDim Preserve Components(1 To LastComponent) As tComponent
    
    With Components(LastComponent)
        
        .X = X
        .Y = Y
        .Component = eComponentType.Label
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
        
        .Text = Text
    End With
    
    AddLabel = LastComponent
    
End Function

Public Function AddShape(ByVal X As Integer, ByVal Y As Integer, _
                            ByVal W As Integer, ByVal H As Integer, _
                            ByRef Color() As Long) As Integer
    
    LastComponent = LastComponent + 1
    
    ReDim Preserve Components(1 To LastComponent) As tComponent
    
    With Components(LastComponent)
        
        .X = X
        .Y = Y
        .W = W
        .H = H
        .Component = eComponentType.Shape
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
    End With
    
    AddShape = LastComponent
    
End Function

Public Function AddTextBox(ByVal X As Integer, ByVal Y As Integer, _
                            ByVal W As Integer, ByVal H As Integer, _
                            ByRef Color() As Long, ByRef ForeColor() As Long, _
                            Optional ByVal ShowOnFocus As Boolean = False) As Integer
    
    LastComponent = LastComponent + 1
    
    ReDim Preserve Components(1 To LastComponent) As tComponent
    
    With Components(LastComponent)
        
        .X = X
        .Y = Y
        .W = W
        .H = H
        .Component = eComponentType.TextBox
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
        
        .ForeColor(0) = ForeColor(0): .ForeColor(1) = ForeColor(1)
        .ForeColor(2) = ForeColor(2): .ForeColor(3) = ForeColor(3)
        
        .IsFocusable = True
        .ShowOnFocus = ShowOnFocus
        
    End With
    
    AddTextBox = LastComponent
    
End Function

Public Sub AppendLine(ByVal ID As Integer, Text As String, TextColor() As Long)
    
    If Not Components(ID).Component = eComponentType.TextArea Then Exit Sub
    
    With Components(ID)
        
        If .LastLine >= MAX_CONSOLE_LINES Then
            Erase .Lines
            .LastLine = 0
        End If
        
        .LastLine = .LastLine + 1
        
        ReDim Preserve .Lines(1 To .LastLine) As TYPE_CONSOLE_LINE
        
        .Lines(.LastLine).Text = Text
        .Lines(.LastLine).Color(0) = TextColor(0)
        .Lines(.LastLine).Color(1) = TextColor(1)
        .Lines(.LastLine).Color(2) = TextColor(2)
        .Lines(.LastLine).Color(3) = TextColor(3)
        
        If .LastLine = 1 Then
            .FirstRender = 1
            .LastRender = 1
        Else
            .LastRender = .LastLine
            
            If .LastLine >= 9 Then
                .FirstRender = .LastLine - 8
            Else
                .FirstRender = 1
            End If
            
        End If
    End With
End Sub

Public Sub AppendLineCC(ByVal ID As Integer, Text As String, _
                        Optional ByVal Red As Integer = 1, Optional ByVal Green As Integer = 1, Optional ByVal Blue As Integer = 1, _
                        Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, _
                        Optional ByVal NewLine As Boolean = True)
                        
    Dim Color(3) As Long
    
    Color(0) = RGB(Red, Green, Blue)
    Color(1) = Color(0)
    Color(2) = Color(0)
    Color(3) = Color(0)
    
    Call AppendLine(ID, Text, Color)
End Sub
Public Sub ClearTextArea(ByVal ID As Integer)

    If Not Components(ID).Component = eComponentType.TextArea Then Exit Sub
    
    With Components(ID)
        
        If .LastLine >= MAX_CONSOLE_LINES Then
            Erase .Lines
            .LastLine = 0
        End If
    End With
End Sub

Public Sub EditLabel(ByVal ID As Integer, Text As String, Color() As Long, Optional ByVal X As Integer = -1, Optional ByVal Y As Integer = -1)

    
    With Components(ID)
        
        If .Component <> eComponentType.Label Then Exit Sub
        
        If X <> -1 Then .X = X
        If Y <> -1 Then .Y = Y
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
        
        .Text = Text
    End With
    
    
End Sub

Public Sub EditShape(ByVal ID As Integer, Color() As Long, _
                        Optional ByVal X As Integer = -1, Optional ByVal Y As Integer = -1, _
                        Optional ByVal W As Integer = -1, Optional ByVal H As Integer = -1)
    
    With Components(ID)
        
        If .Component <> eComponentType.Shape Then Exit Sub
        
        If X <> -1 Then .X = X
        If Y <> -1 Then .Y = Y
        If W <> -1 Then .W = W
        If H <> -1 Then .H = H
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
    End With
    
End Sub

Public Sub EditTextBox(ByVal ID As Integer, Color() As Long, ForeColor() As Long, _
                        Optional ByVal X As Integer = -1, Optional ByVal Y As Integer = -1, _
                        Optional ByVal W As Integer = -1, Optional ByVal H As Integer = -1, _
                        Optional ByVal ShowOnFocus As Boolean = False)
    
    With Components(ID)
        
        If .Component <> eComponentType.TextBox Then Exit Sub
        
        If X <> -1 Then .X = X
        If Y <> -1 Then .Y = Y
        If W <> -1 Then .W = W
        If H <> -1 Then .H = H
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
        
        .ForeColor(0) = ForeColor(0): .ForeColor(1) = ForeColor(1)
        .ForeColor(2) = ForeColor(2): .ForeColor(3) = ForeColor(3)
        
        .IsFocusable = True
        .ShowOnFocus = ShowOnFocus
        
    End With
    
End Sub

Public Function SetComponentFocus(ByVal ID As Integer) As Integer
    
    If Focused <> ID Then
        If Components(ID).IsFocusable Then
            Focused = ID
            SetComponentFocus = ID
        End If
    Else
        SetComponentFocus = ID
    End If
    
End Function

Public Sub RenderComponents(Batch As clsBGFXSpriteBatch)
    
    'todo:
    'Es una lastima el cambio de textura entre el background y las fonts
    
    Dim i As Long
    Dim Component As tComponent
    
    For i = 1 To LastComponent
        Component = Components(i)
        
        With Component
        
            Select Case .Component
            
                Case eComponentType.Label
                    Call Text_Draw(Batch, .X, .Y, .Text, .Color)
                
                Case eComponentType.Shape
                    Call Batch.SetTexture(BackgroundImage.mHandle)
                    Call Batch.Draw(.X, .Y, 0, .W, .H, .Color)
                    
                Case eComponentType.TextBox
                    If .ShowOnFocus Then
                        If Focused = i Then
                            Call Batch.SetTexture(BackgroundImage.mHandle)
                            Call Batch.Draw(.X, .Y, 0, .W, .H, .Color)
                            Call UpdateTextBoxBuffer(Batch, i)
                        End If
                    Else
                        Call Batch.SetTexture(BackgroundImage.mHandle)
                        Call Batch.Draw(.X, .Y, 0, .W, .H, .Color)
                        Call UpdateTextBoxBuffer(Batch, i)
                    End If
                
                Case eComponentType.TextArea
                    Call Batch.SetTexture(BackgroundImage.mHandle)
                    Call Batch.Draw(.X, .Y, 0, .W, .H, .Color())
                    Call UpdateTextArea(Batch, i)
            End Select
            
        End With
    Next
    
End Sub

Private Sub UpdateTextBoxBuffer(Batch As clsBGFXSpriteBatch, ByVal ID As Integer)
    
    If UserWriting Then
        With Components(ID)
            
            If Not StrComp(.TextBuffer, vbNullString) = 0 Then
                Text_Draw Batch, .X + 3, .Y + 3, .TextBuffer, .ForeColor
            End If
            
        End With
    End If
    
End Sub

Private Sub ScrollConsoleUp(ByVal ID As Integer)
    If Components(ID).Component <> eComponentType.TextArea Then Exit Sub
    If Components(ID).LastLine = 0 Then Exit Sub
    
    With Components(ID)
        
        If .FirstRender = 1 Then Exit Sub
        
        .FirstRender = .FirstRender - 1
        .LastRender = .LastRender - 1
        
    End With
    
End Sub

Private Sub ScrollConsoleDown(ByVal ID As Integer)
    If Components(ID).Component <> eComponentType.TextArea Then Exit Sub
    If Components(ID).LastLine = 0 Then Exit Sub
    
    With Components(ID)
        
        If .FirstRender = .LastRender - 8 Then Exit Sub
        
        .FirstRender = .FirstRender + 1
        .LastRender = .LastRender + 1
        
    End With
    
End Sub

Private Sub UpdateTextArea(Batch As clsBGFXSpriteBatch, ByVal ID As Integer)
    
    With Components(ID)
    
        If UBound(.Lines) > 0 Then
            Dim i As Long
            Dim yOffset As Integer
            
            For i = .FirstRender To .LastRender
                Text_Draw Batch, .X + 3, .Y + 2 + yOffset, .Lines(i).Text, .Lines(i).Color
                yOffset = yOffset + 12
            Next
            
        End If
        
    End With
End Sub

Public Sub SetEvents(ByVal ID As Integer, Events As Long)

With Components(ID)

    .HasEvents = True
    
    .EventsPtr = Events
    
End With

End Sub

'@Rezniaq
Public Function Collision(ByVal X As Integer, ByVal Y As Integer) As Integer
 
Dim i                                   As Long
 
'buscamos un objeto que colisione
For i = 1 To LastComponent
    With Components(i)
        'comprobamos X e Y
        If X > .X And X < .X + .W Then
            If Y > .Y And Y < .Y + .H Then
                Collision = i
                Exit Function
            End If
        End If
    End With
Next i
 
'no hay colisión
Collision = -1
 
End Function
 
'@Rezniaq
Public Sub Execute(ByVal ID As Integer, ByVal eventIndex As eComponentEvent, Optional ByVal param3 As Long = 0, Optional ByVal param4 As Long = 0)
 
With Components(ID)
    'si el objeto tiene eventos
    If .HasEvents = True Then
        'si el objeto tiene ESTE evento
        If .EventsPtr <> 0 Then
            'llamamos al sub (un parámetro obligatorio es
            'objectIndex, independientemente de que si el sub
            'lo necesita o no, debe poseerlo como parámetro)
            CallWindowProc .EventsPtr, ID, eventIndex, param3, param4
        End If
    End If
End With
 
End Sub

Public Sub SetFocus(ByVal ID As Integer)
    
    If ID = -1 Then
        Focused = ID
    Else
        If Components(ID).IsFocusable Then Focused = ID
    End If
End Sub

Public Function GetFocused() As Integer

    GetFocused = Focused
    
End Function

Public Function Callback(ByVal param As Long) As Long

        Callback = param
        
End Function

'*********************************************************************************************************************************
'*********************************************************************************************************************************
'*********************************************************************************************************************************
'*****************************************************EVENTS HANDLERS*************************************************************

Public Sub SendTxt_EventHandler(ByVal hWnd As Long, _
                                ByVal msg As Long, _
                                ByVal param3 As Long, _
                                ByVal param4 As Long)

    Dim i       As Long
    Dim tempstr As String
    Static ChatBuffer As String
    
    Select Case msg
    
        Case eComponentEvent.KeyUp

            If param3 = vbKeyReturn Then
                If LenB(Components(hWnd).TextBuffer) <> 0 Then Call ParseUserCommand(Components(hWnd).TextBuffer)
                Components(hWnd).TextBuffer = vbNullString
                ChatBuffer = vbNullString
                UserWriting = False
                
                SetFocus -1
                
            End If
   
        Case eComponentEvent.KeyPress
            If Not (param3 = vbKeyBack) And Not (param3 >= vbKeySpace And param3 <= 250) Then param3 = 0
       
            ChatBuffer = ChatBuffer + ChrW$(param3)
    
            If Len(ChatBuffer) > 160 Then
                Components(hWnd).TextBuffer = "Soy un cheater, avisenle a un gm"
            Else
                'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
            
                For i = 1 To Len(ChatBuffer)
                    param3 = Asc(mid$(ChatBuffer, i, 1))

                    If param3 >= vbKeySpace And param3 <= 250 Then
                        tempstr = tempstr & ChrW$(param3)
                    End If

                    If param3 = vbKeyBack And Len(tempstr) > 0 Then
                        tempstr = Left$(tempstr, Len(tempstr) - 1)
                    End If
                Next i

                If tempstr <> ChatBuffer Then
                    'We only set it if it's different, otherwise the event will be raised
                    'constantly and the client will crush
                    ChatBuffer = tempstr
                End If

                Components(hWnd).TextBuffer = ChatBuffer
            End If
    End Select

End Sub

Public Sub RecTxt_EventHandler(ByVal hWnd As Long, _
                                ByVal msg As Long, _
                                ByVal param3 As Long, _
                                ByVal param4 As Long)

    Select Case msg
    
        Case eComponentEvent.MouseScrollUp
            ScrollConsoleUp hWnd
        Case eComponentEvent.MouseScrollDown
            ScrollConsoleDown hWnd
    End Select
End Sub
