Attribute VB_Name = "mod_Components"
Option Explicit

Private Const MAX_COMBOLIST_LINES As Byte = 5
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
        MouseUp = 7
        MouseDblClick = 8
End Enum

Enum eComponentType
        Label = 0
        TextBox = 1
        Shape = 2
        TextArea = 3
        Rect = 4
        ListBox = 5
        ComboBox = 6
End Enum

Private Type TYPE_CONSOLE_LINE
        Text As String
        Color(3) As Long
End Type

Private Type tComponent 'todo: rehacer, es terrible
        X           As Integer
        Y           As Integer
        w           As Integer
        h           As Integer
        
        Component   As eComponentType
        
        Enable      As Boolean
        Visible     As Boolean
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
        
        SelIndex    As Integer
        
        Expanded    As Boolean 'combobox
        ListID      As Integer 'combobox
        ChildOf     As Integer
        
        PasswChr    As Byte
        
End Type

Private CharHeight      As Integer
Private Focused         As Integer
Private LastComponent   As Integer

Public Components()     As tComponent

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpvDest As Long, ByVal lpvSource As Long, ByVal cbCopy As Long)
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Public Sub InitComponents()
    
    Focused = -1
    
    CharHeight = cfonts(1).CharHeight
    
    If CharHeight = 0 Then
        MsgBox "InitComponents debe colocarse despúes de inicializar los Textos.", vbCritical
        End
    End If
End Sub

Public Sub ClearComponents()
    Erase Components
    Focused = -1
    LastComponent = 0
    
End Sub

Public Function AddListBox(ByVal X As Integer, ByVal Y As Integer, _
                           ByVal w As Integer, ByVal h As Integer, _
                           ByRef BackgroundColor() As Long, Optional ByVal DoRedim As Boolean = True, _
                           Optional ByVal Visible As Boolean = True, Optional ByVal ChildOf As Integer = 0) As Integer
    
    If DoRedim Then
        LastComponent = LastComponent + 1
    
        ReDim Preserve Components(1 To LastComponent) As tComponent
    End If
    
    With Components(LastComponent)
    
        .X = X: .w = w
        .Y = Y: .h = h
        
        .Component = eComponentType.ListBox
        
        .Color(0) = BackgroundColor(0): .Color(1) = BackgroundColor(1)
        .Color(2) = BackgroundColor(2): .Color(3) = BackgroundColor(3)
        
        .Visible = Visible
        
        .Enable = True
        
        .SelIndex = -1
        .ChildOf = ChildOf
    End With
    
    Call SetEvents(LastComponent, Callback(AddressOf ListBox_EventHandler))
    
    AddListBox = LastComponent
    
End Function

Public Function AddComboBox(ByVal X As Integer, ByVal Y As Integer, _
                            ByVal w As Integer, _
                            ByRef BackgroundColor() As Long) As Integer
                            
    LastComponent = LastComponent + 2
    
    ReDim Preserve Components(1 To LastComponent) As tComponent
    
    With Components(LastComponent - 1)
    
        .X = X: .w = w
        .Y = Y: .h = CharHeight
        
        .Component = eComponentType.ComboBox
        
        .Color(0) = BackgroundColor(0): .Color(1) = BackgroundColor(1)
        .Color(2) = BackgroundColor(2): .Color(3) = BackgroundColor(3)
            
        .Enable = True
        .Visible = True
        
        .ListID = AddListBox(X + w, Y, w, 0, BackgroundColor, False, False, LastComponent - 1)
    End With
    
    Call SetEvents(LastComponent - 1, Callback(AddressOf ComboBox_EventHandler))
    AddComboBox = (LastComponent - 1)
    
End Function

Public Function AddRect(ByVal X As Integer, ByVal Y As Integer, _
                        ByVal w As Integer, ByVal h As Integer) As Integer
                            
                
    LastComponent = LastComponent + 1
    ReDim Preserve Components(1 To LastComponent) As tComponent
    
    With Components(LastComponent)
    
        .X = X: .w = w
        .Y = Y: .h = h
        
        .Component = eComponentType.Rect
        
        .Enable = True
        .Visible = True
    End With
    
    AddRect = LastComponent
    
End Function

Public Function AddTextArea(ByVal X As Integer, ByVal Y As Integer, _
                            ByVal w As Integer, ByVal h As Integer, _
                            Color() As Long) As Integer
    
    LastComponent = LastComponent + 1
    
    ReDim Preserve Components(1 To LastComponent) As tComponent
        
    With Components(LastComponent)
    
        .X = X: .w = w
        .Y = Y: .h = h
        
        .Component = eComponentType.TextArea
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
        
        .Enable = True
        .Visible = True
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
        
        .Visible = True
    End With
    
    AddLabel = LastComponent
    
End Function

Public Function AddShape(ByVal X As Integer, ByVal Y As Integer, _
                            ByVal w As Integer, ByVal h As Integer, _
                            ByRef Color() As Long) As Integer
    
    LastComponent = LastComponent + 1
    
    ReDim Preserve Components(1 To LastComponent) As tComponent
    
    
    With Components(LastComponent)
        
        .X = X
        .Y = Y
        .w = w
        .h = h
        .Component = eComponentType.Shape
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
        
        .Enable = True
        .Visible = True
    End With
    
    AddShape = LastComponent
    
End Function

Public Function AddTextBox(ByVal X As Integer, ByVal Y As Integer, _
                            ByVal w As Integer, ByVal h As Integer, _
                            ByRef Color() As Long, ByRef ForeColor() As Long, _
                            Optional ByVal ShowOnFocus As Boolean = False, Optional ByVal PasswChr As Boolean = False) As Integer
    
    LastComponent = LastComponent + 1
    
    ReDim Preserve Components(1 To LastComponent) As tComponent
    
    With Components(LastComponent)
        
        .X = X
        .Y = Y
        .w = w
        .h = h
        .Component = eComponentType.TextBox
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
        
        .ForeColor(0) = ForeColor(0): .ForeColor(1) = ForeColor(1)
        .ForeColor(2) = ForeColor(2): .ForeColor(3) = ForeColor(3)
        
        .IsFocusable = True
        .ShowOnFocus = ShowOnFocus
        
        .PasswChr = PasswChr
        
        .Enable = True
        .Visible = True
    End With
    
    Call SetEvents(LastComponent, Callback(AddressOf TextBox_EventHandler))
    
    AddTextBox = LastComponent
    
End Function

Public Sub TabComponent()
    
    Dim i As Long
    Dim startId As Long
    
    If LastComponent <> 0 Then
        
        If Focused <> -1 Then
            i = Focused
            startId = Focused
        End If
        
        i = i + 1
        
        Do Until ((Components(i).IsFocusable = True And Components(i).Visible) Or startId = Focused)
        
            If LastComponent = i Then
                i = 0
            End If
            
            i = i + 1
        Loop
        
        Focused = i
    End If
End Sub

'Listboxex and combos
Public Sub InsertText(ByVal ID As Integer, Text As String, TextColor() As Long)
    
    If Not (Components(ID).Component = eComponentType.ComboBox Or Components(ID).Component = eComponentType.ListBox) Then Exit Sub
    
    Dim i As Integer
    
    If Components(ID).Component = eComponentType.ComboBox Then i = Components(ID).ListID Else i = ID
    
    With Components(i)
        
        .LastLine = .LastLine + 1
    
        If .LastLine - 1 = 0 Then
            ReDim .Lines(1 To .LastLine) As TYPE_CONSOLE_LINE 'reused :^)
        Else
            ReDim Preserve .Lines(1 To .LastLine) As TYPE_CONSOLE_LINE
        End If
        
        .Lines(.LastLine).Text = Text
        .Lines(.LastLine).Color(0) = TextColor(0)
        .Lines(.LastLine).Color(1) = TextColor(1)
        .Lines(.LastLine).Color(2) = TextColor(2)
        .Lines(.LastLine).Color(3) = TextColor(3)
        
        Dim LastDrawableLine As Integer
        
        If Components(ID).Component = eComponentType.ComboBox Then .h = .h + CharHeight + 1
        
        LastDrawableLine = Fix(.h / CharHeight)
        
        If .LastLine = 1 Then
            .FirstRender = 1
            .LastRender = 1
            If Components(ID).Component = eComponentType.ComboBox Then Components(ID).Text = Text
        Else
            .LastRender = .LastLine
            
            If .LastLine >= LastDrawableLine Then
                .FirstRender = .LastLine - (LastDrawableLine - 1)
            Else
                .FirstRender = 1
            End If
        End If
        
    End With
End Sub

Public Sub AppendLine(ByVal ID As Integer, Text As String, TextColor() As Long)
    
    If Not Components(ID).Component = eComponentType.TextArea Then Exit Sub
    
    With Components(ID)
        
        If .LastLine >= MAX_CONSOLE_LINES Then
            .LastLine = 0
        End If
        
        .LastLine = .LastLine + 1
        
        If .LastLine - 1 = 0 Then
            ReDim .Lines(1 To .LastLine) As TYPE_CONSOLE_LINE
        Else
            ReDim Preserve .Lines(1 To .LastLine) As TYPE_CONSOLE_LINE
        End If
        
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
                        Optional ByVal Red As Integer = 1, Optional ByVal Green As Integer = 1, Optional ByVal blue As Integer = 1, _
                        Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, _
                        Optional ByVal NewLine As Boolean = True)
                        
    Dim Color(3) As Long
    
    Color(0) = RGB(Red, Green, blue)
    Color(1) = Color(0)
    Color(2) = Color(0)
    Color(3) = Color(0)
    
    Call AppendLine(ID, Text, Color)
End Sub

Public Sub ClearTextArea(ByVal ID As Integer, Optional ByVal Forced As Boolean = False)

    If Not Components(ID).Component = eComponentType.TextArea Then Exit Sub
    
    With Components(ID)
        
        If (.LastLine >= MAX_CONSOLE_LINES Or Forced) Then
            .LastLine = 0
            .FirstRender = 0
            .LastRender = 0
            
            ReDim .Lines(1) As TYPE_CONSOLE_LINE
            
        End If
    End With
End Sub

Public Function GetComboText(ByVal ID As Integer) As String
    GetComboText = Components(ID).Text
End Function

Public Function GetSelectedValue(ByVal ID As Integer) As String
    If Components(ID).SelIndex <> 0 Then _
        GetSelectedValue = Components(ID).Lines(Components(ID).SelIndex).Text
End Function

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
                        Optional ByVal w As Integer = -1, Optional ByVal h As Integer = -1)
    
    With Components(ID)
        
        If .Component <> eComponentType.Shape Then Exit Sub
        
        If X <> -1 Then .X = X
        If Y <> -1 Then .Y = Y
        If w <> -1 Then .w = w
        If h <> -1 Then .h = h
        
        .Color(0) = Color(0): .Color(1) = Color(1)
        .Color(2) = Color(2): .Color(3) = Color(3)
    End With
    
End Sub

Public Sub EditTextBox(ByVal ID As Integer, Color() As Long, ForeColor() As Long, _
                        Optional ByVal X As Integer = -1, Optional ByVal Y As Integer = -1, _
                        Optional ByVal w As Integer = -1, Optional ByVal h As Integer = -1, _
                        Optional ByVal ShowOnFocus As Boolean = False)
    
    With Components(ID)
        
        If .Component <> eComponentType.TextBox Then Exit Sub
        
        If X <> -1 Then .X = X
        If Y <> -1 Then .Y = Y
        If w <> -1 Then .w = w
        If h <> -1 Then .h = h
        
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

Public Sub RenderComponents()
    
    Dim i As Long
    Dim Component As tComponent
    
    For i = 1 To LastComponent
        Component = Components(i)
        
        With Component
            
            If .Visible = False Then GoTo NextLoop
            
            Select Case .Component
            
                Case eComponentType.Label
                    Call Text_Draw(.X, .Y, .Text, .Color)
                
                Case eComponentType.Shape
                    Call Draw_Box(.X, .Y, .w, .h, .Color)
                    
                Case eComponentType.TextBox
                    If .ShowOnFocus Then
                        If Focused = i Then
                            Call Draw_Box(.X, .Y, .w, .h, .Color)
                            Call UpdateTextBoxBuffer(i)
                        End If
                    Else
                        Call Draw_Box(.X, .Y, .w, .h, .Color)
                        Call UpdateTextBoxBuffer(i)
                    End If
                
                Case eComponentType.TextArea
                    Call Draw_Box(.X, .Y, .w, .h, .Color)
                    Call UpdateTextArea(i)
                
                Case eComponentType.ComboBox
                    Call Draw_Box(.X, .Y, .w, .h, .Color)
                    Call Text_Draw(.X + 3, .Y - 1, .Text, White)
                    Call Draw_Box(.X + .w - 10, .Y, .h, .h, Gray)
                    
                    If .Expanded Then
                        Call Text_Draw(.X + .w - 8, .Y - 1, "<", Black)
                    Else
                        Call Text_Draw(.X + .w - 8, .Y - 1, ">", Black)
                    End If
                    
                Case eComponentType.ListBox
                    Call DrawListBox(i)
            End Select
            
        End With
        
NextLoop:
    Next
    
End Sub

Private Sub DrawListBox(ByVal ID As Integer)
    
    With Components(ID)
    
        Dim i As Long
        Dim yOffset As Integer
        
        Call Draw_Box(.X, .Y, .w, .h, .Color)
        
        If .FirstRender = 0 Then Exit Sub
        
        For i = .FirstRender To .LastRender
            If i = .SelIndex Then
                Call Draw_Box(.X, .Y + 1 + yOffset, .w, CharHeight, Gray)
                Call Text_Draw(.X + 3, .Y + 1 + yOffset, .Lines(i).Text, .Lines(i).Color)
            Else
                Call Text_Draw(.X + 3, .Y + 1 + yOffset, .Lines(i).Text, .Lines(i).Color)
            End If
            yOffset = yOffset + CharHeight
        Next
        
    End With
End Sub

Private Sub UpdateTextBoxBuffer(ByVal ID As Integer)
    
    'If UserWriting Then
        With Components(ID)
            
            If Not StrComp(.TextBuffer, vbNullString) = 0 Then
                
                Dim renderstr As String
                If .PasswChr Then
                    renderstr = String$(Len(.TextBuffer), "*")
                Else
                    renderstr = .TextBuffer
                End If
                
                If Focused = ID Then
                    Call Text_Draw(.X + 3, .Y + 3, renderstr + "|", .ForeColor)
                Else
                    Call Text_Draw(.X + 3, .Y + 3, renderstr, .ForeColor)
                End If
            Else
                If Focused = ID Then
                    Call Text_Draw(.X + 3, .Y + 3, "|", .ForeColor)
                End If
                
            End If
            
        End With
    'End If
    
End Sub

Private Sub ScrollListUp(ByVal ID As Integer)
    If Components(ID).Component <> eComponentType.ListBox Then Exit Sub
    If Components(ID).LastLine = 0 Then Exit Sub
    
    With Components(ID)
        
        If .FirstRender = 1 Then Exit Sub
        
        .FirstRender = .FirstRender - 1
        .LastRender = .LastRender - 1
        
    End With
    
End Sub

Private Sub ScrollListDown(ByVal ID As Integer)
    If Components(ID).Component <> eComponentType.ListBox Then Exit Sub
    If Components(ID).LastLine = 0 Then Exit Sub
    
    With Components(ID)
        
        Dim LastDrawableLine As Integer
        
        LastDrawableLine = Fix(.h / CharHeight)
        
        If .LastLine = (LastDrawableLine + .FirstRender) - 1 Then Exit Sub
        
        .FirstRender = .FirstRender + 1
        .LastRender = .LastRender + 1
        
    End With
    
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

Private Sub UpdateTextArea(ByVal ID As Integer)
    
    With Components(ID)
    
        Dim i As Long
        Dim yOffset As Integer
            
        For i = .FirstRender To .LastRender
            Text_Draw .X + 3, .Y + 2 + yOffset, .Lines(i).Text, .Lines(i).Color
            yOffset = yOffset + 12
        Next
        
    End With
End Sub

Public Sub SetEvents(ByVal ID As Integer, Events As Long)

With Components(ID)

    .HasEvents = True
    
    .EventsPtr = Events
    
End With

End Sub

Public Function GetComponentText(ByVal ID As Integer) As String
        
    GetComponentText = Components(ID).TextBuffer
    
End Function

'@Rezniaq
Public Function Collision(ByVal X As Integer, ByVal Y As Integer) As Integer
 
Dim i                                   As Long
 
'buscamos un objeto que colisione
For i = 1 To LastComponent
    With Components(i)
        'comprobamos X e Y
        If X > .X And X < .X + .w Then
            If Y > .Y And Y < .Y + .h Then
                If .Visible And .Enable Then
                    Collision = i
                    Exit Function
                End If
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
    If .Enable Then
        If .HasEvents = True Then
            'si el objeto tiene ESTE evento
            If .EventsPtr <> 0 Then
                'llamamos al sub (un parámetro obligatorio es
                'objectIndex, independientemente de que si el sub
                'lo necesita o no, debe poseerlo como parámetro)
                CallWindowProc .EventsPtr, ID, eventIndex, param3, param4
            End If
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

Public Sub HideComponents(ParamArray Comps() As Variant)
    Dim i As Long
    
    For i = 0 To UBound(Comps)
        Components(Comps(i)).Visible = False
    Next
    
End Sub

Public Sub ShowComponents(ParamArray Comps() As Variant)
    Dim i As Long
    
    For i = 0 To UBound(Comps)
        Components(Comps(i)).Visible = True
    Next
    
End Sub

Public Sub DisableComponents(ParamArray Comps() As Variant)
    Dim i As Long
    
    For i = 0 To UBound(Comps)
        Components(Comps(i)).Enable = False
    Next
End Sub

Public Sub EnableComponents(ParamArray Comps() As Variant)
    Dim i As Long
    
    For i = 0 To UBound(Comps)
        Components(Comps(i)).Enable = True
    Next
End Sub

Public Sub SetChild(ByVal FatherID As Integer, ByVal ChildID As Integer) 'no siblings yet
    Components(ChildID).ChildOf = FatherID
End Sub

'*********************************************************************************************************************************
'*********************************************************************************************************************************
'*********************************************************************************************************************************
'*****************************************************EVENTS HANDLERS*************************************************************

'This Override the primitive method TextBox_EventHandler
Public Sub txtRepPass_EventHandler(ByVal hwnd As Long, _
                                   ByVal msg As Long, _
                                   ByVal param3 As Long, _
                                   ByVal param4 As Long)
    Dim i As Long
    Dim tempstr As String
    Dim Buffer As String
    
    Buffer = Components(hwnd).TextBuffer
    
    With Components(hwnd)
    
        Select Case msg
            
            Case eComponentEvent.MouseUp
                Call SetFocus(hwnd)
                
            Case eComponentEvent.KeyPress
                If Not (param3 = vbKeyBack) And Not (param3 >= vbKeySpace And param3 <= 250) Then param3 = 0
                
    
                Buffer = Buffer + ChrW$(param3)
                
                'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
                For i = 1 To Len(Buffer)
                    param3 = Asc(mid$(Buffer, i, 1))
    
                    If param3 >= vbKeySpace And param3 <= 250 Then
                        tempstr = tempstr & ChrW$(param3)
                    End If
    
                    If param3 = vbKeyBack And Len(tempstr) > 0 Then
                        tempstr = Left$(tempstr, Len(tempstr) - 1)
                    End If
                Next i
    
                If tempstr <> Buffer Then
                    'We only set it if it's different, otherwise the event will be raised
                    'constantly and the client will crush
                    Buffer = tempstr
                End If
    
                Components(hwnd).TextBuffer = Buffer

                If StrComp(.TextBuffer, Components(.ChildOf).TextBuffer) = 0 Then
                
                    .ForeColor(0) = Green(0): .ForeColor(1) = Green(1)
                    .ForeColor(2) = Green(2): .ForeColor(3) = Green(3)
                    
                Else
                    .ForeColor(0) = Red(0): .ForeColor(1) = Red(1)
                    .ForeColor(2) = Red(2): .ForeColor(3) = Red(3)
                End If
        End Select
        
    End With
End Sub


Public Sub btnLogin_EventHandler(ByVal hwnd As Long, _
                                ByVal msg As Long, _
                                ByVal param3 As Long, _
                                ByVal param4 As Long)
    
    Select Case msg
    
        Case eComponentEvent.MouseUp
            frmConnect.LoginUser
    End Select
End Sub

Public Sub btnNewCharacter_EventHandler(ByVal hwnd As Long, _
                                ByVal msg As Long, _
                                ByVal param3 As Long, _
                                ByVal param4 As Long)

    EstadoLogin = E_MODO.Dados
        
    If frmMain.Socket1.Connected Then
        frmMain.Socket1.Disconnect
        frmMain.Socket1.Cleanup
        DoEvents
    End If
            
    frmMain.Socket1.HostName = CurServerIP
    frmMain.Socket1.RemotePort = CurServerPort
    frmMain.Socket1.Connect

End Sub

Public Sub ComboBox_EventHandler(ByVal hwnd As Long, _
                                 ByVal msg As Long, _
                                 ByVal param3 As Long, _
                                 ByVal param4 As Long)
    
    If msg <> 0 Then
        
        Select Case msg
            
            Case eComponentEvent.MouseDown
                Components(hwnd).Expanded = Not Components(hwnd).Expanded
                Components(Components(hwnd).ListID).Visible = Components(hwnd).Expanded
                
            Case eComponentEvent.MouseScrollUp: If Components(hwnd).Expanded Then Call ScrollListUp(hwnd)
            Case eComponentEvent.MouseScrollDown: If Components(hwnd).Expanded Then Call ScrollListDown(hwnd)
            
            
        End Select
    End If
End Sub

Public Sub btnSiguiente_EventHandler(ByVal hwnd As Long, _
                                     ByVal msg As Long, _
                                     ByVal param3 As Long, _
                                     ByVal param4 As Long)
    
    Select Case msg
        
        Case eComponentEvent.MouseUp
            
            If GetRenderState() = eRenderState.eLogin Then Exit Sub
            
            If GetRenderState = eRenderState.eNewCharSkills Then
                Call frmConnect.LoginNewChar
            Else
                Call ChangeRenderState(GetRenderState() + 1)
            End If
    End Select
    
End Sub

Public Sub btnAtras_EventHandler(ByVal hwnd As Long, _
                                     ByVal msg As Long, _
                                     ByVal param3 As Long, _
                                     ByVal param4 As Long)
    
    Select Case msg
        
        Case eComponentEvent.MouseUp
            
            If GetRenderState() = eRenderState.eLogin Then Exit Sub
            
            If GetRenderState = eRenderState.eNewCharInfo Then
                Call frmConnect.CloseNewChar
            Else
                Call ChangeRenderState(GetRenderState() - 1)
            End If
    End Select
    
End Sub

'Primitive Events

Private Sub ListBox_EventHandler(ByVal hwnd As Long, _
                                 ByVal msg As Long, _
                                 ByVal param3 As Long, _
                                 ByVal param4 As Long)

    Dim X As Integer, Y As Integer
    Dim LastDrawableLine As Integer
        
        
    Call LongToIntegers(param3, X, Y)
    
    If msg <> 0 Then
        
        Y = Y - (Components(hwnd).Y - CharHeight \ 2)
        Y = (Y + 2) \ CharHeight
        
        With Components(hwnd)
        
            Select Case msg
                
                Case eComponentEvent.MouseDown
                    LastDrawableLine = Fix(.h / CharHeight)
                    
                    If Y > 0 And Y <= LastDrawableLine Then
                        .SelIndex = (.FirstRender - 1) + Y
                        
                        Dim cho As Integer
                        cho = .ChildOf
                        
                        If cho <> 0 Then
                            
                            Components(cho).Text = .Lines(.SelIndex).Text
                        End If
                        
                    End If
                    
                Case eComponentEvent.MouseScrollUp: Call ScrollListUp(hwnd)
                Case eComponentEvent.MouseScrollDown: Call ScrollListDown(hwnd)
    
            End Select
        End With
    End If
    
End Sub

Public Sub TextBox_EventHandler(ByVal hwnd As Long, _
                                ByVal msg As Long, _
                                ByVal param3 As Long, _
                                ByVal param4 As Long)

    Dim i As Long
    Dim tempstr As String
    Dim Buffer As String
    
    Buffer = Components(hwnd).TextBuffer
    
    Select Case msg
        
        Case eComponentEvent.MouseUp
            Call SetFocus(hwnd)
            
        Case eComponentEvent.KeyPress
            If Not (param3 = vbKeyBack) And Not (param3 >= vbKeySpace And param3 <= 250) Then param3 = 0
            

            Buffer = Buffer + ChrW$(param3)
            
            'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
            For i = 1 To Len(Buffer)
                param3 = Asc(mid$(Buffer, i, 1))

                If param3 >= vbKeySpace And param3 <= 250 Then
                    tempstr = tempstr & ChrW$(param3)
                End If

                If param3 = vbKeyBack And Len(tempstr) > 0 Then
                    tempstr = Left$(tempstr, Len(tempstr) - 1)
                End If
            Next i

            If tempstr <> Buffer Then
                'We only set it if it's different, otherwise the event will be raised
                'constantly and the client will crush
                Buffer = tempstr
            End If

            Components(hwnd).TextBuffer = Buffer
    End Select
    
End Sub
