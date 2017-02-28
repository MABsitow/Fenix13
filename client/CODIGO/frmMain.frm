VERSION 5.00
Object = "{33101C00-75C3-11CF-A8A0-444553540000}#1.0#0"; "CSWSK32.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   11520
   ClientLeft      =   360
   ClientTop       =   300
   ClientWidth     =   15360
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   768
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1024
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin SocketWrenchCtrl.Socket Socket1 
      Left            =   1560
      Top             =   0
      _Version        =   65536
      _ExtentX        =   741
      _ExtentY        =   741
      _StockProps     =   0
      AutoResolve     =   0   'False
      Backlog         =   1
      Binary          =   -1  'True
      Blocking        =   0   'False
      Broadcast       =   0   'False
      BufferSize      =   10240
      HostAddress     =   ""
      HostFile        =   ""
      HostName        =   ""
      InLine          =   0   'False
      Interval        =   0
      KeepAlive       =   0   'False
      Library         =   ""
      Linger          =   0
      LocalPort       =   0
      LocalService    =   ""
      Protocol        =   0
      RemotePort      =   0
      RemoteService   =   ""
      ReuseAddress    =   0   'False
      Route           =   -1  'True
      Timeout         =   10000
      Type            =   1
      Urgent          =   0   'False
   End
   Begin VB.Timer macrotrabajo 
      Enabled         =   0   'False
      Left            =   1080
      Top             =   0
   End
   Begin VB.Timer TrainingMacro 
      Enabled         =   0   'False
      Interval        =   3121
      Left            =   720
      Top             =   0
   End
   Begin VB.Timer Macro 
      Interval        =   750
      Left            =   360
      Top             =   0
   End
   Begin VB.Timer Second 
      Enabled         =   0   'False
      Interval        =   1050
      Left            =   0
      Top             =   0
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   15000
      Top             =   0
      Width           =   375
   End
   Begin VB.Shape MainViewShp 
      BorderColor     =   &H00000080&
      Height          =   8160
      Left            =   120
      Top             =   2280
      Visible         =   0   'False
      Width           =   11040
   End
   Begin VB.Menu mnuObj 
      Caption         =   "Objeto"
      Visible         =   0   'False
      Begin VB.Menu mnuTirar 
         Caption         =   "Tirar"
      End
      Begin VB.Menu mnuUsar 
         Caption         =   "Usar"
      End
      Begin VB.Menu mnuEquipar 
         Caption         =   "Equipar"
      End
   End
   Begin VB.Menu mnuNpc 
      Caption         =   "NPC"
      Visible         =   0   'False
      Begin VB.Menu mnuNpcDesc 
         Caption         =   "Descripcion"
      End
      Begin VB.Menu mnuNpcComerciar 
         Caption         =   "Comerciar"
         Visible         =   0   'False
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************COMPONENTS********************
Public lblLvl           As Integer          'store the id of the component
Public GldLvl           As Integer
Public lblExp           As Integer
Public lblPorcExp       As Integer
Public lblMode          As Integer
Public SendTxt          As Integer
Public lblClase         As Integer
Public lblFaccion       As Integer
Public lblRecompensa    As Integer
Public lblArmor         As Integer
Public lblShielder      As Integer  'shielder? wtf ofi
Public lblWeapon        As Integer
Public lblHelm          As Integer
Public lblName          As Integer
Public RecTxt           As Integer

Public PicInv           As Integer

'***************COMPONENTS********************


Private isDebug As Boolean

Public TX As Byte
Public TY As Byte
Public MouseX As Long
Public MouseY As Long
Public FormMouseX As Long
Public FormMouseY As Long
Public MouseBoton As Long
Public MouseShift As Long
Private clicX As Long
Private clicY As Long

Public IsPlaying As Byte

Private clsFormulario As clsFormMovementManager

Public picSkillStar As Picture

Dim PuedeMacrear As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    Dim ID As Integer
    ID = mod_Components.GetFocused()
    
    If ID <> -1 Then
        Call mod_Components.Execute(ID, eComponentEvent.KeyPress, KeyAscii)
    End If
        
End Sub

Private Sub Form_Load()
    
    If NoRes Then
        ' Handles Form movement (drag and drop).
        Set clsFormulario = New clsFormMovementManager
        clsFormulario.Initialize Me, 120
    End If
    
'    InvEqu.Picture = LoadPicture(DirGraficos & "CentroInventario.jpg")
    
    InitComponents
    
    Hook_Main Me.hWnd
    
    Me.Left = 0
    Me.Top = 0
End Sub

Private Sub InitComponents()
    
    lblName = AddLabel("", 816, 24, White())
    lblLvl = AddLabel("Nivel: ", 847, 76, White())
    GldLvl = AddLabel(0, 960, 418, White())
    lblExp = AddLabel("Exp: 999999999/99999999", 847, 100, White())
    lblPorcExp = AddLabel("33.33%", 847, 88, Cyan())
    lblMode = AddLabel("1 Normal", 16, 128, White())
    
    SendTxt = AddTextBox(74, 122, 428, 21, Black(), White(), True)
    
    Call SetEvents(SendTxt, Callback(AddressOf SendTxt_EventHandler))
    
    lblClase = AddLabel("C", 944, 80, White())
    lblFaccion = AddLabel("F", 960, 80, White())
    lblRecompensa = AddLabel("R", 976, 80, White())
    lblArmor = AddLabel("00/00", 78, 744, Red())
    lblShielder = AddLabel("00/00", 342, 744, Red())
    lblWeapon = AddLabel("00/00", 464, 744, Red())
    lblHelm = AddLabel("00/00", 196, 744, Red())
    
    RecTxt = AddTextArea(14, 14, 734, 119, Black)
    
    Call SetEvents(RecTxt, Callback(AddressOf RecTxt_EventHandler))
    
    AppendLine RecTxt, "Esto es un motd", White()
    AppendLine RecTxt, "Hola", Red()
    
    PicInv = AddRect(800, 223, 160, 160)
    
    Call SetEvents(PicInv, Callback(AddressOf PicInv_EventHandler))
    
End Sub

Public Sub LightSkillStar(ByVal bTurnOn As Boolean)
    'If bTurnOn Then
    '    imgAsignarSkill.Picture = picSkillStar
    'Else
    '    Set imgAsignarSkill.Picture = Nothing
    'End If
End Sub

Private Sub cmdMoverHechi_Click(Index As Integer)
'    If hlst.Visible = True Then
'        If hlst.ListIndex = -1 Then Exit Sub
'        Dim sTemp As String
'
'        Select Case Index
'            Case 1 'subir
'                If hlst.ListIndex = 0 Then Exit Sub
'            Case 0 'bajar
'                If hlst.ListIndex = hlst.ListCount - 1 Then Exit Sub
'        End Select
'
'        Call WriteMoveSpell(Index = 1, hlst.ListIndex + 1)
'
'        Select Case Index
'            Case 1 'subir
'                sTemp = hlst.List(hlst.ListIndex - 1)
'                hlst.List(hlst.ListIndex - 1) = hlst.List(hlst.ListIndex)
'                hlst.List(hlst.ListIndex) = sTemp
'                hlst.ListIndex = hlst.ListIndex - 1
'            Case 0 'bajar
'                sTemp = hlst.List(hlst.ListIndex + 1)
'                hlst.List(hlst.ListIndex + 1) = hlst.List(hlst.ListIndex)
'                hlst.List(hlst.ListIndex) = sTemp
'                hlst.ListIndex = hlst.ListIndex + 1
'        End Select
'    End If
End Sub

Public Sub ActivarMacroHechizos()
'    If Not hlst.Visible Then
'        Call AddtoRichTextBox(frmMain.RecTxt, "Debes tener seleccionado el hechizo para activar el auto-lanzar", 0, 200, 200, False, True, True)
'        Exit Sub
'    End If
'
    TrainingMacro.Interval = INT_MACRO_HECHIS
    TrainingMacro.Enabled = True
'    Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos activado", 0, 200, 200, False, True, True)
    'Call ControlSM(eSMType.mSpells, True)
End Sub

Public Sub DesactivarMacroHechizos()
    TrainingMacro.Enabled = False
    'Call AddtoRichTextBox(frmMain.RecTxt, "Auto lanzar hechizos desactivado", 0, 150, 150, False, True, True)
    'Call ControlSM(eSMType.mSpells, False)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
'***************************************************
'Autor: Unknown
'Last Modification: 18/11/2009
'18/11/2009: ZaMa - Ahora se pueden poner comandos en los mensajes personalizados (execpto guildchat y privados)
'***************************************************

    'If (Not SendTxt.Visible) And (Not SendCMSTXT.Visible) Then
        
    If Not mod_Components.GetFocused() = SendTxt Then
        'Checks if the key is valid
        If LenB(CustomKeys.ReadableName(KeyCode)) > 0 Then
            Select Case KeyCode
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleMusic)
                    'audio.MusicActivated = Not 'audio.MusicActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleSound)
                    'audio.SoundActivated = Not 'audio.SoundActivated
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleFxs)
                    'audio.SoundEffectsActivated = Not 'audio.SoundEffectsActivated
                
                Case CustomKeys.BindedKey(eKeyType.mKeyGetObject)
                    Call AgarrarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyEquipObject)
                    Call EquiparItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyToggleNames)
                    Nombres = Not Nombres
                
                Case CustomKeys.BindedKey(eKeyType.mKeyTamAnimal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .Bold, .Italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Domar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeySteal)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .Bold, .Italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Robar)
                    End If
                    
                Case CustomKeys.BindedKey(eKeyType.mKeyHide)
                    If UserEstado = 1 Then
                        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .Bold, .Italic)
                        End With
                    Else
                        Call WriteWork(eSkill.Ocultarse)
                    End If
                                    
                Case CustomKeys.BindedKey(eKeyType.mKeyDropObject)
                    Call TirarItem
                
                Case CustomKeys.BindedKey(eKeyType.mKeyUseObject)
                    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
                        
                    If MainTimer.Check(TimersIndex.UseItemWithU) Then
                        Call UsarItem
                    End If
                
                Case CustomKeys.BindedKey(eKeyType.mKeyRequestRefresh)
                    If MainTimer.Check(TimersIndex.SendRPU) Then
                        Call WriteRequestPositionUpdate
                        Beep
                    End If
            End Select
        Else
            Select Case KeyCode
                'Custom messages!
                Case vbKey0 To vbKey9
                    Dim CustomMessage As String
                    
                    CustomMessage = CustomMessages.Message((KeyCode - 39) Mod 10)
                    If LenB(CustomMessage) <> 0 Then
                        ' No se pueden mandar mensajes personalizados de clan o privado!
                        If UCase(Left(CustomMessage, 5)) <> "/CMSG" And _
                            Left(CustomMessage, 1) <> "\" Then
                            
                            Call ParseUserCommand(CustomMessage)
                        End If
                    End If
            End Select
        End If
    End If
    
    Select Case KeyCode
    
        Case vbKeyF1:
            Call EditLabel(lblMode, "1 Normal", White())
            TalkMode = 1
            ChoosingWhisper = False
            MousePointer = 1
            
        Case vbKeyF2:
            'Call AddtoRichTextBox(frmMain.RecTxt, "Has click sobre el usuario al que quieres susurrar.", 255, 255, 255, 1, 0)
            Call EditLabel(lblMode, "2 Susurrar", White())
            MousePointer = 2
            TalkMode = 2
            ChoosingWhisper = True
            
        Case vbKeyF3:
            Call EditLabel(lblMode, "3 Clan", White())
            TalkMode = 3
            ChoosingWhisper = False
            MousePointer = 1
            
        Case vbKeyF4:
            Call EditLabel(lblMode, "4 Grito", White())
            TalkMode = 4
            ChoosingWhisper = False
            MousePointer = 1
            
        Case vbKeyF5:
            Call EditLabel(lblMode, "5 Rol", White())
            TalkMode = 5
            ChoosingWhisper = False
            MousePointer = 1
            
        Case vbKeyF6:
            Call EditLabel(lblMode, "6 Party", White())
            TalkMode = 6
            ChoosingWhisper = False
            MousePointer = 1
        
        Case vbKeyF12:
            If isDebug Then
                Call Video.SetDebug(DEBUG_MODE_NONE)
                isDebug = False
            Else
                Call Video.SetDebug(DEBUG_MODE_STATS)
                isDebug = True
            End If
                
      '  Case CustomKeys.BindedKey(eKeyType.mKeyShowOptions)
       '     Call frmOpciones.Show(vbModeless, frmMain)
        
        Case CustomKeys.BindedKey(eKeyType.mKeyMeditate)
            If UserMinMAN = UserMaxMAN Then Exit Sub
            
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .Bold, .Italic)
                End With
                Exit Sub
            End If
                
            If Not PuedeMacrear Then
                'AddtoRichTextBox frmMain.RecTxt, "No tan rápido..!", 255, 255, 255, False, False, True
            Else
                Call WriteMeditate
                PuedeMacrear = False
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyCastSpellMacro)
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .Bold, .Italic)
                End With
                Exit Sub
            End If
            
            If TrainingMacro.Enabled Then
                DesactivarMacroHechizos
            Else
                ActivarMacroHechizos
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyWorkMacro)
            If UserEstado = 1 Then
                With FontTypes(FontTypeNames.FONTTYPE_INFO)
                    Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .Bold, .Italic)
                End With
                Exit Sub
            End If
            
            If macrotrabajo.Enabled Then
                Call DesactivarMacroTrabajo
            Else
                Call ActivarMacroTrabajo
            End If
        
        Case CustomKeys.BindedKey(eKeyType.mKeyExitGame)
            If frmMain.macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
            Call WriteQuit
            
        Case CustomKeys.BindedKey(eKeyType.mKeyAttack)
            If Shift <> 0 Then Exit Sub
            
            If Not MainTimer.Check(TimersIndex.Arrows, False) Then Exit Sub 'Check if arrows interval has finished.
            If Not MainTimer.Check(TimersIndex.CastSpell, False) Then 'Check if spells interval has finished.
                If Not MainTimer.Check(TimersIndex.CastAttack) Then Exit Sub 'Corto intervalo Golpe-Hechizo
            Else
                If Not MainTimer.Check(TimersIndex.Attack) Or UserDescansar Or UserMeditar Then Exit Sub
            End If
            
            If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
            If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
            Call WriteAttack
        
        Case CustomKeys.BindedKey(eKeyType.mKeyTalk)
            'If SendCMSTXT.Visible Then Exit Sub
            If mod_Components.GetFocused() = SendTxt Then
                Call mod_Components.Execute(SendTxt, eComponentEvent.KeyUp, KeyCode)
            Else
            
                If (Not Comerciando) And (Not MirandoAsignarSkills) And _
                  (Not frmMSG.Visible) And (Not MirandoForo) And _
                  (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
                    Call mod_Components.SetFocus(SendTxt)
                    UserWriting = True
                End If
            End If
    End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    MouseBoton = Button
    MouseShift = Shift
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    clicX = X
    clicY = Y
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    
    UnHook_Main Me.hWnd
    
    If prgRun = True Then
        prgRun = False
        Cancel = 1
    End If
End Sub

Private Sub imgAsignarSkill_Click()
    Dim i As Integer
    
    LlegaronSkills = False
    Call WriteRequestSkills
    Call FlushBuffer
    
    Do While Not LlegaronSkills
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    LlegaronSkills = False
    
    For i = 1 To NUMSKILLS
        frmSkills3.Text1(i).Caption = UserSkills(i)
    Next i
    
    Alocados = SkillPoints
    frmSkills3.puntos.Caption = SkillPoints
    frmSkills3.Show , frmMain

End Sub

Private Sub imgClanes_Click()
Call MsgBox("Sistema deshabilitado.", vbInformation, "Argentum Online")
'If frmGuildLeader.Visible Then Unload frmGuildLeader
 '   Call WriteRequestGuildLeaderInfo
End Sub

Private Sub imgEstadisticas_Click()
    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
    Call WriteRequestAtributes
    Call WriteRequestSkills
    Call WriteRequestMiniStats
    Call WriteRequestFame
    Call FlushBuffer
    Do While Not LlegaronSkills Or Not LlegaronAtrib Or Not LlegoFama
        DoEvents 'esperamos a que lleguen y mantenemos la interfaz viva
    Loop
    frmEstadisticas.Iniciar_Labels
    frmEstadisticas.Show , frmMain
    LlegaronAtrib = False
    LlegaronSkills = False
    LlegoFama = False
End Sub

Private Sub imgGrupo_Click()
Call MsgBox("Sistema deshabilitado.", vbInformation, "Argentum Online")
'    Call WriteRequestPartyForm
End Sub

Private Sub imgInvScrollDown_Click()
    'Call Inventario.ScrollInventory(True)
End Sub

Private Sub imgInvScrollUp_Click()
    'Call Inventario.ScrollInventory(False)
End Sub

Private Sub imgMapa_Click()
    Call frmMapa.Show(vbModeless, frmMain)
End Sub

Private Sub imgOpciones_Click()
Call MsgBox("Funciones deshabilitadas.", vbInformation, "Argentum Online")
'Call frmOpciones.Show(vbModeless, frmMain)
End Sub

Private Sub lblScroll_Click(Index As Integer)
    'Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub lblClase_Click()
    Call WriteRequestClaseForm
End Sub

Private Sub lblCerrar_Click()
    prgRun = False
End Sub

Private Sub lblFaccion_Click()
    Call WriteRequestFaccionForm
End Sub

Private Sub lblMinimizar_Click()
    Me.WindowState = 1
End Sub

Private Sub lblRecompensa_Click()
    Call WriteRequestRecompensaForm
End Sub


Private Sub Image1_Click()
ParseUserCommand "/salir"
End Sub

Private Sub Macro_Timer()
    PuedeMacrear = True
End Sub

Private Sub macrotrabajo_Timer()
    If Inventario.SelectedItem = 0 Then
        Call DesactivarMacroTrabajo
        Exit Sub
    End If
    
    'Macros are disabled if not using Argentum!
    If Not Application.IsAppActive() Then
        Call DesactivarMacroTrabajo
        Exit Sub
    End If
    
    If UsingSkill = eSkill.Pesca Or UsingSkill = eSkill.Talar Or UsingSkill = eSkill.Mineria Or _
                UsingSkill = FundirMetal Or (UsingSkill = eSkill.Herreria And Not frmHerrero.Visible) Then
        Call WriteWorkLeftClick(TX, TY, UsingSkill)
        UsingSkill = 0
    End If
    
    'If Inventario.OBJType(Inventario.SelectedItem) = eObjType.otWeapon Then
     If Not (frmCarp.Visible = True) Then Call UsarItem
End Sub

Public Sub ActivarMacroTrabajo()
    macrotrabajo.Interval = INT_MACRO_TRABAJO
    macrotrabajo.Enabled = True
    'Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo ACTIVADO", 0, 200, 200, False, True, True)
End Sub

Public Sub DesactivarMacroTrabajo()
    macrotrabajo.Enabled = False
    MacroBltIndex = 0
    UsingSkill = 0
    MousePointer = vbDefault
    'Call AddtoRichTextBox(frmMain.RecTxt, "Macro Trabajo DESACTIVADO", 0, 200, 200, False, True, True)
End Sub


Private Sub mnuEquipar_Click()
    Call EquiparItem
End Sub

Private Sub mnuNPCComerciar_Click()
    Call WriteLeftClick(TX, TY)
    Call WriteCommerceStart
End Sub

Private Sub mnuNpcDesc_Click()
    Call WriteLeftClick(TX, TY)
End Sub

Private Sub mnuTirar_Click()
    Call TirarItem
End Sub

Private Sub mnuUsar_Click()
    Call UsarItem
End Sub

Private Sub Coord_Click()
    'Call AddtoRichTextBox(frmMain.RecTxt, "Estas coordenadas son tu ubicación en el mapa. Utiliza la letra L para corregirla si esta no se corresponde con la del servidor por efecto del Lag.", 255, 255, 255, False, False, True)
End Sub

Private Sub SendTxt_KeyUp(KeyCode As Integer, Shift As Integer)
    'Send text
    If KeyCode = vbKeyReturn Then
        If LenB(stxtbuffer) <> 0 Then Call ParseUserCommand(stxtbuffer)
        
        stxtbuffer = ""
        'SendTxt.Text = ""
        KeyCode = 0
        'SendTxt.Visible = False
        
        'If PicInv.Visible Then
        '    PicInv.SetFocus
        'Else
        '    hlst.SetFocus
        'End If
    End If
End Sub

Private Sub Second_Timer()
'    If Not DialogosClanes Is Nothing Then DialogosClanes.PassTimer
End Sub

'[END]'

''''''''''''''''''''''''''''''''''''''
'     ITEM CONTROL                   '
''''''''''''''''''''''''''''''''''''''

Private Sub TirarItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .Bold, .Italic)
        End With
    Else
        If (Inventario.SelectedItem > 0 And Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Or (Inventario.SelectedItem = FLAGORO) Then
            If Inventario.Amount(Inventario.SelectedItem) = 1 Then
                Call WriteDrop(Inventario.SelectedItem, 1)
            Else
                If Inventario.Amount(Inventario.SelectedItem) > 1 Then
                    If Not Comerciando Then frmCantidad.Show , frmMain
                End If
            End If
        End If
    End If
End Sub

Private Sub AgarrarItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
            Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .Bold, .Italic)
        End With
    Else
        Call WritePickUp
    End If
End Sub

Private Sub UsarItem()
    If pausa Then Exit Sub
    
    If Comerciando Then Exit Sub
    
    If TrainingMacro.Enabled Then DesactivarMacroHechizos
    
    If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteUseItem(Inventario.SelectedItem)
End Sub

Private Sub EquiparItem()
    If UserEstado = 1 Then
        With FontTypes(FontTypeNames.FONTTYPE_INFO)
                Call ShowConsoleMsg("¡¡Estás muerto!!", .Red, .Green, .Blue, .Bold, .Italic)
        End With
    Else
        If Comerciando Then Exit Sub
        
        If (Inventario.SelectedItem > 0) And (Inventario.SelectedItem < MAX_INVENTORY_SLOTS + 1) Then _
        Call WriteEquipItem(Inventario.SelectedItem)
    End If
End Sub



''''''''''''''''''''''''''''''''''''''
'     HECHIZOS CONTROL               '
''''''''''''''''''''''''''''''''''''''

Private Sub TrainingMacro_Timer()
    'If Not hlst.Visible Then
    '    DesactivarMacroHechizos
    '    Exit Sub
    'End If
    
    'Macros are disabled if focus is not on Argentum!
    If Not Application.IsAppActive() Then
        DesactivarMacroHechizos
        Exit Sub
    End If
    
    If Comerciando Then Exit Sub
    
    'If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.CastSpell, False) Then
    '    Call WriteCastSpell(hlst.ListIndex + 1)
    '    Call WriteWork(eSkill.Magia)
    'End If
    
    Call ConvertCPtoTP(MouseX, MouseY, TX, TY)
    
    If UsingSkill = Magia And Not MainTimer.Check(TimersIndex.CastSpell) Then Exit Sub
    
    If UsingSkill = Proyectiles And Not MainTimer.Check(TimersIndex.Attack) Then Exit Sub
    
    Call WriteWorkLeftClick(TX, TY, UsingSkill)
    UsingSkill = 0
End Sub

Private Sub cmdLanzar_Click()
'    If hlst.List(hlst.ListIndex) <> "(None)" And MainTimer.Check(TimersIndex.Work, False) Then
'        If UserEstado = 1 Then
'            With FontTypes(FontTypeNames.FONTTYPE_INFO)
'                Call ShowConsoleMsg("¡¡Estás muerto!!", .red, .green, .blue, .bold, .italic)
'            End With
'        Else
'            Call WriteCastSpell(hlst.ListIndex + 1)
'            Call WriteWork(eSkill.Magia)
'            UsaMacro = True
'        End If
'    End If
End Sub

Private Sub CmdLanzar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    UsaMacro = False
    CnTd = 0
End Sub

Private Sub cmdINFO_Click()
'    If hlst.ListIndex <> -1 Then
'        Call WriteSpellInfo(hlst.ListIndex + 1)
'    End If
End Sub

Private Sub DespInv_Click(Index As Integer)
    'Inventario.ScrollInventory (Index = 0)
End Sub

Private Sub Form_Click()
    If Cartel Then Cartel = False

    If Not Comerciando Then
        Call ConvertCPtoTP(MouseX, MouseY, TX, TY)
        
        If Not InGameArea() Then Exit Sub
        
            If MouseBoton <> vbRightButton Then
                If ChoosingWhisper Then
                    ChoosingWhisper = False
                    MousePointer = 1
                    Dim SelChar As Integer
                    SelChar = MapData(TX, TY).CharIndex
                    
                    If SelChar <> UserCharIndex And SelChar > 0 Then
                        WhisperTarget = charlist(SelChar).Nombre
                    End If
                    
                End If
                
                '[ybarra]
                If UsaMacro Then
                    CnTd = CnTd + 1
                    If CnTd = 3 Then
                        Call WriteUseSpellMacro
                        CnTd = 0
                    End If
                    UsaMacro = False
                End If
                '[/ybarra]
                If UsingSkill = 0 Then
                    Call WriteLeftClick(TX, TY)
                Else
                
                    If TrainingMacro.Enabled Then Call DesactivarMacroHechizos
                    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
                    
                    If Not MainTimer.Check(TimersIndex.Arrows, False) Then 'Check if arrows interval has finished.
                        frmMain.MousePointer = vbDefault
                        UsingSkill = 0
                        With FontTypes(FontTypeNames.FONTTYPE_TALK)
                            'Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .red, .green, .blue, .bold, .italic)
                        End With
                        Exit Sub
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Proyectiles Then
                        If Not MainTimer.Check(TimersIndex.Arrows) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                'Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar proyectiles tan rápido.", .red, .green, .blue, .bold, .italic)
                            End With
                            Exit Sub
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If UsingSkill = Magia Then
                        If Not MainTimer.Check(TimersIndex.Attack, False) Then 'Check if attack interval has finished.
                            If Not MainTimer.Check(TimersIndex.CastAttack) Then 'Corto intervalo de Golpe-Magia
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    'Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rápido.", .red, .green, .blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        Else
                            If Not MainTimer.Check(TimersIndex.CastSpell) Then 'Check if spells interval has finished.
                                frmMain.MousePointer = vbDefault
                                UsingSkill = 0
                                With FontTypes(FontTypeNames.FONTTYPE_TALK)
                                    'Call AddtoRichTextBox(frmMain.RecTxt, "No puedes lanzar hechizos tan rapido.", .red, .green, .blue, .bold, .italic)
                                End With
                                Exit Sub
                            End If
                        End If
                    End If
                    
                    'Splitted because VB isn't lazy!
                    If (UsingSkill = Pesca Or UsingSkill = Robar Or UsingSkill = Talar Or UsingSkill = Mineria Or UsingSkill = FundirMetal) Then
                        If Not MainTimer.Check(TimersIndex.Work) Then
                            frmMain.MousePointer = vbDefault
                            UsingSkill = 0
                            Exit Sub
                        End If
                    End If
                    
                    If frmMain.MousePointer <> 2 Then Exit Sub 'Parcheo porque a veces tira el hechizo sin tener el cursor (NicoNZ)
                    
                    frmMain.MousePointer = vbDefault
                    Call WriteWorkLeftClick(TX, TY, UsingSkill)
                    UsingSkill = 0
                End If
            End If
            
            If MouseBoton = vbRightButton Then
                Call WriteWarpChar("YO", UserMap, TX, TY)
            End If
          
    End If
End Sub

Private Sub Form_DblClick()
'**************************************************************
'Author: Unknown
'Last Modify Date: 12/27/2007
'12/28/2007: ByVal - Chequea que la ventana de comercio y boveda no este abierta al hacer doble clic a un comerciante, sobrecarga la lista de items.
'**************************************************************
    If Not MirandoForo And Not Comerciando Then 'frmComerciar.Visible And Not frmBancoObj.Visible Then
        Call WriteDoubleClick(TX, TY)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'Dim Col As Integer
    
    'Col = Collision(X, Y)
    
    FormMouseX = X
    FormMouseY = Y
    
    MouseX = X - MainViewShp.Left
    MouseY = Y - MainViewShp.Top
    
    'Trim to fit screen
    If MouseX < 0 Then
        MouseX = 0
    ElseIf MouseX > MainViewShp.Width Then
        MouseX = MainViewShp.Width
    End If
    
    'Trim to fit screen
    If MouseY < 0 Then
        MouseY = 0
    ElseIf MouseY > MainViewShp.Height Then
        MouseY = MainViewShp.Height
    End If
        
    'If Col <> -1 Then
    '    Call mod_Components.Execute(Col, eComponentEvent.MouseMove)
    'End If
End Sub

Private Sub hlst_KeyDown(KeyCode As Integer, Shift As Integer)
       KeyCode = 0
End Sub

Private Sub hlst_KeyPress(KeyAscii As Integer)
       KeyAscii = 0
End Sub

Private Sub hlst_KeyUp(KeyCode As Integer, Shift As Integer)
        KeyCode = 0
End Sub

Private Sub lblDropGold_Click()

    Inventario.SelectGold
    If UserGLD > 0 Then
        If Not Comerciando Then frmCantidad.Show , frmMain
    End If
    
End Sub

Private Sub Label4_Click()
    'call 'audio.PlayWave(SND_CLICK)

'    InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centroinventario.jpg")
'
'    ' Activo controles de inventario
'    PicInv.Visible = True
'    imgInvScrollUp.Visible = True
'    imgInvScrollDown.Visible = True
'
'    ' Desactivo controles de hechizo
'    hlst.Visible = False
'    cmdINFO.Visible = False
'    CmdLanzar.Visible = False
'
'    cmdMoverHechi(0).Visible = False
'    cmdMoverHechi(1).Visible = False
    
End Sub

Private Sub Label7_Click()
    'call 'audio.PlayWave(SND_CLICK)

'    InvEqu.Picture = LoadPicture(App.path & "\Graficos\Centrohechizos.jpg")
'
'    ' Activo controles de hechizos
'    hlst.Visible = True
'    cmdINFO.Visible = True
'    CmdLanzar.Visible = True
'
'    cmdMoverHechi(0).Visible = True
'    cmdMoverHechi(1).Visible = True
'
'    ' Desactivo controles de inventario
'    PicInv.Visible = False
'    imgInvScrollUp.Visible = False
'    imgInvScrollDown.Visible = False

End Sub

Private Sub picInv_DblClick()
    If frmCarp.Visible Or frmHerrero.Visible Then Exit Sub
    
    If Not MainTimer.Check(TimersIndex.UseItemWithDblClick) Then Exit Sub
    
    If macrotrabajo.Enabled Then Call DesactivarMacroTrabajo
    
    Call UsarItem
End Sub

Private Sub picInv_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'call 'audio.PlayWave(SND_CLICK)
End Sub

Private Sub RecTxt_Change()
'On Error Resume Next  'el .SetFocus causaba errores al salir y volver a entrar
'    If Not Application.IsAppActive() Then Exit Sub
'
'    If SendTxt.Visible Then
'        SendTxt.SetFocus
'    ElseIf Me.SendCMSTXT.Visible Then
'        SendCMSTXT.SetFocus
'    ElseIf (Not Comerciando) And (Not MirandoAsignarSkills) And _
'        (Not frmMSG.Visible) And (Not MirandoForo) And _
'        (Not frmEstadisticas.Visible) And (Not frmCantidad.Visible) Then
'
'        If PicInv.Visible Then
'            PicInv.SetFocus
'        ElseIf hlst.Visible Then
'            hlst.SetFocus
'        End If
'    End If
End Sub

Private Sub RecTxt_KeyDown(KeyCode As Integer, Shift As Integer)
'    If PicInv.Visible Then
'        PicInv.SetFocus
'    Else
'        hlst.SetFocus
'    End If
End Sub

Private Sub SendTxt_Change()
'**************************************************************
'Author: Unknown
'Last Modify Date: 3/06/2006
'3/06/2006: Maraxus - impedí se inserten caractéres no imprimibles
'**************************************************************
'    If Len(SendTxt.Text) > 160 Then
'        stxtbuffer = "Soy un cheater, avisenle a un gm"
'    Else
'        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
'        Dim i As Long
'        Dim tempstr As String
'        Dim CharAscii As Integer
'
'        For i = 1 To Len(SendTxt.Text)
'            CharAscii = Asc(mid$(SendTxt.Text, i, 1))
'            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
'                tempstr = tempstr & Chr$(CharAscii)
'            End If
'        Next i
'
'        If tempstr <> SendTxt.Text Then
'            'We only set it if it's different, otherwise the event will be raised
'            'constantly and the client will crush
'            SendTxt.Text = tempstr
'        End If
'
'        stxtbuffer = SendTxt.Text
'    End If
End Sub

Private Sub SendTxt_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendCMSTXT_KeyUp(KeyCode As Integer, Shift As Integer)
'    'Send text
'    If KeyCode = vbKeyReturn Then
'        'Say
'        If stxtbuffercmsg <> "" Then
'            Call ParseUserCommand("/CMSG " & stxtbuffercmsg)
'        End If
'
'        stxtbuffercmsg = ""
'        SendCMSTXT.Text = ""
'        KeyCode = 0
'        Me.SendCMSTXT.Visible = False
'
'        If PicInv.Visible Then
'            PicInv.SetFocus
'        Else
'            hlst.SetFocus
'        End If
'    End If
End Sub

Private Sub SendCMSTXT_KeyPress(KeyAscii As Integer)
    If Not (KeyAscii = vbKeyBack) And _
       Not (KeyAscii >= vbKeySpace And KeyAscii <= 250) Then _
        KeyAscii = 0
End Sub

Private Sub SendCMSTXT_Change()
'    If Len(SendCMSTXT.Text) > 160 Then
'        stxtbuffercmsg = "Soy un cheater, avisenle a un GM"
'    Else
'        'Make sure only valid chars are inserted (with Shift + Insert they can paste illegal chars)
'        Dim i As Long
'        Dim tempstr As String
'        Dim CharAscii As Integer
'
'        For i = 1 To Len(SendCMSTXT.Text)
'            CharAscii = Asc(mid$(SendCMSTXT.Text, i, 1))
'            If CharAscii >= vbKeySpace And CharAscii <= 250 Then
'                tempstr = tempstr & Chr$(CharAscii)
'            End If
'        Next i
'
'        If tempstr <> SendCMSTXT.Text Then
'            'We only set it if it's different, otherwise the event will be raised
'            'constantly and the client will crush
'            SendCMSTXT.Text = tempstr
'        End If
'
'        stxtbuffercmsg = SendCMSTXT.Text
'    End If
End Sub


''''''''''''''''''''''''''''''''''''''
'     SOCKET1                        '
''''''''''''''''''''''''''''''''''''''
Private Sub Socket1_Connect()
    
    'Clean input and output buffers
    Call incomingData.Clear
    Call outgoingData.Clear

    Second.Enabled = True

    Select Case EstadoLogin
        Case E_MODO.CrearNuevoPj
            Call Login
        
        Case E_MODO.Normal
           Call Login
        
        Case E_MODO.Dados
            'call 'audio.PlayMIDI("7.mid")
            frmCrearPersonaje.Show vbModal
    End Select
End Sub

Private Sub Socket1_Disconnect()
    Dim i As Long
    
    Second.Enabled = False
    Connected = False
    
    Socket1.Cleanup
    
    frmConnect.MousePointer = vbNormal
    
    Do While i < Forms.Count - 1
        i = i + 1
        
        If Forms(i).Name <> Me.Name And Forms(i).Name <> frmConnect.Name And Forms(i).Name <> frmCrearPersonaje.Name Then
            Unload Forms(i)
        End If
    Loop
    
    On Local Error GoTo 0
    
    If Not frmCrearPersonaje.Visible Then
        frmConnect.Visible = True
    End If
    
    frmMain.Visible = False
    
    pausa = False
    UserMeditar = False
    
    UserClase = 0
    UserSexo = 0
    UserRaza = 0
    UserHogar = 0
    UserEmail = ""
    
    For i = 1 To NUMSKILLS
        UserSkills(i) = 0
    Next i

    For i = 1 To NUMATRIBUTOS
        UserAtributos(i) = 0
    Next i
    
    For i = 1 To MAX_INVENTORY_SLOTS
        
    Next i
    
    macrotrabajo.Enabled = False

    SkillPoints = 0
    Alocados = 0
End Sub

Private Sub Socket1_LastError(ErrorCode As Integer, ErrorString As String, Response As Integer)
    '*********************************************
    'Handle socket errors
    '*********************************************
    If ErrorCode = 24036 Then
        Call MsgBox("Por favor espere, intentando completar conexion.", vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
        Exit Sub
    End If
    
    Call MsgBox(ErrorString, vbApplicationModal + vbInformation + vbOKOnly + vbDefaultButton1, "Error")
    frmConnect.MousePointer = 1
    Response = 0
    Second.Enabled = False

    frmMain.Socket1.Disconnect
    
    If Not frmCrearPersonaje.Visible Then
        frmConnect.Show
    Else
        frmCrearPersonaje.MousePointer = 0
    End If
End Sub

Private Sub Socket1_Read(dataLength As Integer, IsUrgent As Integer)
    Dim RD As String
    Dim Data() As Byte
    
    Call Socket1.Read(RD, dataLength)
    Data = StrConv(RD, vbFromUnicode)
    
    If RD = vbNullString Then Exit Sub

    'Put data in the buffer
    Call incomingData.Wrap(Data)
    
    'Send buffer to Handle data
    Call HandleIncomingData
End Sub

Private Function InGameArea() As Boolean
'***************************************************
'Author: NicoNZ
'Last Modification: 04/07/08
'Checks if last click was performed within or outside the game area.
'***************************************************
    If clicX < MainViewShp.Left Or clicX > MainViewShp.Left + MainViewShp.Width Then Exit Function
    If clicY < MainViewShp.Top Or clicY > MainViewShp.Top + MainViewShp.Height Then Exit Function
    
    InGameArea = True
End Function
