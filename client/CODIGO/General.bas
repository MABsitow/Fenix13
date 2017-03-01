Attribute VB_Name = "Mod_General"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 M�rquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Mat�as Fernando Peque�o
'
'This program is free software; you can redistribute it and/or modify
'it under the terms of the Affero General Public License;
'either version 1 of the License, or any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'Affero General Public License for more details.
'
'You should have received a copy of the Affero General Public License
'along with this program; if not, you can find it at http://www.affero.org/oagpl.html
'
'Argentum Online is based on Baronsoft's VB6 Online RPG
'You can contact the original creator of ORE at aaron@baronsoft.com
'for more information about ORE please visit http://www.baronsoft.com/
'
'
'You can contact me at:
'morgolock@speedy.com.ar
'www.geocities.com/gmorgolock
'Calle 3 n�mero 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'C�digo Postal 1900
'Pablo Ignacio M�rquez

Option Explicit

Public iplst As String

Public bFogata As Boolean

Public bLluvia() As Byte ' Array para determinar si
'debemos mostrar la animacion de la lluvia

Private lFrameTimer As Long

Public Function DirGraficos() As String
    DirGraficos = App.path & "\" & Config_Inicio.DirGraficos & "\"
End Function

Public Function DirSound() As String
    DirSound = App.path & "\" & Config_Inicio.DirSonidos & "\"
End Function

Public Function DirMidi() As String
    DirMidi = App.path & "\" & Config_Inicio.DirMusica & "\"
End Function

Public Function DirMapas() As String
    DirMapas = App.path & "\" & Config_Inicio.DirMapas & "\"
End Function

Public Function DirExtras() As String
    DirExtras = App.path & "\EXTRAS\"
End Function

Public Function RandomNumber(ByVal LowerBound As Long, ByVal UpperBound As Long) As Long
    'Initialize randomizer
    Randomize Timer
    
    'Generate random number
    RandomNumber = (UpperBound - LowerBound) * Rnd + LowerBound
End Function

Public Function GetRawName(ByRef sName As String) As String
'***************************************************
'Author: ZaMa
'Last Modify Date: 13/01/2010
'Last Modified By: -
'Returns the char name without the clan name (if it has it).
'***************************************************

    Dim Pos As Integer
    
    Pos = InStr(1, sName, "<")
    
    If Pos > 0 Then
        GetRawName = Trim(Left(sName, Pos - 1))
    Else
        GetRawName = sName
    End If

End Function

Sub CargarAnimArmas()
On Error Resume Next

    Dim loopC As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopC = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopC).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopC, "Dir1")), 0
        InitGrh WeaponAnimData(loopC).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopC, "Dir2")), 0
        InitGrh WeaponAnimData(loopC).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopC, "Dir3")), 0
        InitGrh WeaponAnimData(loopC).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopC, "Dir4")), 0
    Next loopC
End Sub

Sub CargarColores()
On Error Resume Next
    Dim archivoC As String
    
    archivoC = App.path & "\init\colores.dat"
    
    If Not FileExist(archivoC, vbArchive) Then
'TODO : Si hay que reinstalar, porque no cierra???
        Call MsgBox("ERROR: no se ha podido cargar los colores. Falta el archivo colores.dat, reinstale el juego", vbCritical + vbOKOnly)
        Exit Sub
    End If
    
    Dim i As Long
    
    For i = 0 To 46 '49 y 50 reservados para ciudadano y criminal
        ColoresPJ(i) = RGB(CInt(GetVar(archivoC, CStr(i), "R")), CInt(GetVar(archivoC, CStr(i), "G")), CInt(GetVar(archivoC, CStr(i), "B")))
    Next i
    
    ' Crimi
    ColoresPJ(50) = RGB(CInt(GetVar(archivoC, "CR", "R")), _
    CInt(GetVar(archivoC, "CR", "G")), _
    CInt(GetVar(archivoC, "CR", "B")))
    
    ' Ciuda
    ColoresPJ(49) = RGB(CInt(GetVar(archivoC, "CI", "R")), _
    CInt(GetVar(archivoC, "CI", "G")), _
    CInt(GetVar(archivoC, "CI", "B")))
    
    ' Neutral
    ColoresPJ(48) = RGB(CInt(GetVar(archivoC, "NE", "R")), _
    CInt(GetVar(archivoC, "NE", "G")), _
    CInt(GetVar(archivoC, "NE", "B")))
    
    ColoresPJ(47) = RGB(CInt(GetVar(archivoC, "NW", "R")), _
    CInt(GetVar(archivoC, "NW", "G")), _
    CInt(GetVar(archivoC, "NW", "B")))
    
End Sub

Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopC As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopC = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopC).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopC, "Dir1")), 0
        InitGrh ShieldAnimData(loopC).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopC, "Dir2")), 0
        InitGrh ShieldAnimData(loopC).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopC, "Dir3")), 0
        InitGrh ShieldAnimData(loopC).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopC, "Dir4")), 0
    Next loopC
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal Red As Integer = -1, Optional ByVal Green As Integer, Optional ByVal Blue As Integer, Optional ByVal Bold As Boolean = False, Optional ByVal Italic As Boolean = False, Optional ByVal bCrLf As Boolean = True)
'******************************************
'Adds text to a Richtext box at the bottom.
'Automatically scrolls to new text.
'Text box MUST be multiline and have a 3D
'apperance!
'Pablo (ToxicWaste) 01/26/2007 : Now the list refeshes properly.
'Juan Mart�n Sotuyo Dodero (Maraxus) 03/29/2007 : Replaced ToxicWaste's code for extra performance.
'******************************************r
    With RichTextBox
        If Len(.Text) > 1000 Then
            'Get rid of first line
            .SelStart = InStr(1, .Text, vbCrLf) + 1
            .SelLength = Len(.Text) - .SelStart + 2
            .TextRTF = .SelRTF
        End If
        
        .SelStart = Len(.Text)
        .SelLength = 0
        .SelBold = Bold
        .SelItalic = Italic
        
        If Not Red = -1 Then .SelColor = RGB(Red, Green, Blue)
        
        If bCrLf And Len(.Text) > 0 Then Text = vbCrLf & Text
        .SelText = Text
        
        RichTextBox.Refresh
    End With
End Sub

'TODO : Never was sure this is really necessary....
'TODO : 08/03/2006 - (AlejoLp) Esto hay que volarlo...
Public Sub RefreshAllChars()
'*****************************************************************
'Goes through the charlist and replots all the characters on the map
'Used to make sure everyone is visible
'*****************************************************************
    Dim loopC As Long
    
    For loopC = 1 To LastChar
        If charlist(loopC).Active = 1 Then
            MapData(charlist(loopC).Pos.X, charlist(loopC).Pos.Y).CharIndex = loopC
        End If
    Next loopC
End Sub

Sub SaveGameini()
    'Grabamos los datos del usuario en el Game.ini
    Config_Inicio.Name = "BetaTester"
    Config_Inicio.Password = "DammLamers"
    Config_Inicio.Puerto = UserPort
    
    Call EscribirGameIni(Config_Inicio)
End Sub

Function AsciiValidos(ByVal cad As String) As Boolean
    Dim car As Byte
    Dim i As Long
    
    cad = LCase$(cad)
    
    For i = 1 To Len(cad)
        car = Asc(mid$(cad, i, 1))
        
        If ((car < 97 Or car > 122) Or car = Asc("�")) And (car <> 255) And (car <> 32) Then
            Exit Function
        End If
    Next i
    
    AsciiValidos = True
End Function

Function CheckUserData(ByVal checkemail As Boolean) As Boolean
    'Validamos los datos del user
    Dim loopC As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Direcci�n de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopC = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopC, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inv�lido. El caract�r " & Chr$(CharAscii) & " no est� permitido.")
            Exit Function
        End If
    Next loopC
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopC = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, loopC, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inv�lido. El caract�r " & Chr$(CharAscii) & " no est� permitido.")
            Exit Function
        End If
    Next loopC
    
    CheckUserData = True
End Function

Sub UnloadAllForms()
On Error Resume Next

    Dim mifrm As Form
    
    For Each mifrm In Forms
        Unload mifrm
    Next
End Sub

Function LegalCharacter(ByVal KeyAscii As Integer) As Boolean
'*****************************************************************
'Only allow characters that are Win 95 filename compatible
'*****************************************************************
    'if backspace allow
    If KeyAscii = 8 Then
        LegalCharacter = True
        Exit Function
    End If
    
    'Only allow space, numbers, letters and special characters
    If KeyAscii < 32 Or KeyAscii = 44 Then
        Exit Function
    End If
    
    If KeyAscii > 126 Then
        Exit Function
    End If
    
    'Check for bad special characters in between
    If KeyAscii = 34 Or KeyAscii = 42 Or KeyAscii = 47 Or KeyAscii = 58 Or KeyAscii = 60 Or KeyAscii = 62 Or KeyAscii = 63 Or KeyAscii = 92 Or KeyAscii = 124 Then
        Exit Function
    End If
    
    'else everything is cool
    LegalCharacter = True
End Function

Sub SetConnected()
'*****************************************************************
'Sets the client to "Connect" mode
'*****************************************************************
    'Set Connected
    Connected = True
    
    Call SaveGameini
    
    'Unload the connect form
    Unload frmCrearPersonaje
    Unload frmConnect
    
    'todo
    'frmMain.lblName.Caption = UserName
    'Load main form
    frmMain.Visible = True
        
    FPSFLAG = True
End Sub

Sub CargarTip()
    Dim N As Integer
    N = RandomNumber(1, UBound(Tips))
    
    frmtip.tip.Caption = Tips(N)
End Sub

Sub MoveTo(ByVal Direccion As E_Heading)
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/28/2008
'Last Modified By: Lucas Tavolaro Ortiz (Tavo)
' 06/03/2006: AlejoLp - Elimine las funciones Move[NSWE] y las converti a esta
' 12/08/2007: Tavo    - Si el usuario esta paralizado no se puede mover.
' 06/28/2008: NicoNZ - Saqu� lo que imped�a que si el usuario estaba paralizado se ejecute el sub.
'***************************************************
    Dim LegalOk As Boolean
    
    If Cartel Then Cartel = False
    
    Select Case Direccion
        Case E_Heading.NORTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y - 1)
        Case E_Heading.EAST
            LegalOk = MoveToLegalPos(UserPos.X + 1, UserPos.Y)
        Case E_Heading.SOUTH
            LegalOk = MoveToLegalPos(UserPos.X, UserPos.Y + 1)
        Case E_Heading.WEST
            LegalOk = MoveToLegalPos(UserPos.X - 1, UserPos.Y)
    End Select
    
    If LegalOk And Not UserParalizado Then
        Call WriteWalk(Direccion)
        If Not UserDescansar And Not UserMeditar Then
            MoveCharbyHead UserCharIndex, Direccion
            MoveScreen Direccion
            'MoveCamera Direccion
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call WriteChangeHeading(Direccion)
        End If
    End If
    
    If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
    
    ' Update 3D sounds!
    'call 'audio.MoveListener(UserPos.X, UserPos.Y)
End Sub

Sub RandomMove()
'***************************************************
'Author: Alejandro Santos (AlejoLp)
'Last Modify Date: 06/03/2006
' 06/03/2006: AlejoLp - Ahora utiliza la funcion MoveTo
'***************************************************
    Call MoveTo(RandomNumber(NORTH, WEST))
End Sub

Private Sub CheckKeys()
'*****************************************************************
'Checks keys and respond
'*****************************************************************
    Static LastMovement As Long
    
    'No input allowed while Argentum is not the active window
    If Not Application.IsAppActive() Then Exit Sub
    
    'No walking when in commerce or banking.
    If Comerciando Then Exit Sub
    
    'No walking while writting in the forum.
    If MirandoForo Then Exit Sub
    
    'If game is paused, abort movement.
    If pausa Then Exit Sub
    
    'TODO: Deber�a informarle por consola?
    If Traveling Then Exit Sub

    'Control movement interval (this enforces the 1 step loss when meditating / resting client-side)
    If GetTickCount - LastMovement > 56 Then
        LastMovement = GetTickCount
    Else
        Exit Sub
    End If
    
    'Don't allow any these keys during movement..
    If UserMoving = 0 Then
        If Not UserEstupido Then
            'Move Up
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(NORTH)
                'todo
                'frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
                Exit Sub
            End If
            
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(EAST)
                'frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.y & ")"
                'frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(SOUTH)
                'frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(WEST)
                'frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
                Exit Sub
            End If
            
            ' We haven't moved - Update 3D sounds!
            'call 'audio.MoveListener(UserPos.X, UserPos.Y)
        Else
            Dim kp As Boolean
            kp = (GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyUp)) < 0) Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Or _
                GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0
            
            If kp Then
                Call RandomMove
            Else
                ' We haven't moved - Update 3D sounds!
                'call 'audio.MoveListener(UserPos.X, UserPos.Y)
            End If
            
            If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
            'frmMain.Coord.Caption = "(" & UserPos.x & "," & UserPos.y & ")"
            'frmMain.Coord.Caption = "X: " & UserPos.X & " Y: " & UserPos.Y
        End If
    End If
End Sub

'CSEH: ErrLog
Sub SwitchMap(ByVal Map As Integer)
    Dim Y As Long
    Dim X As Long
    Dim Handle As Integer
    Dim Reader As New CsBuffer
    Dim Data() As Byte
    Dim ByFlags As Byte
        
    Handle = FreeFile()
    
    Open DirMapas & "Mapa" & Map & ".mcl" For Binary As Handle
        Seek Handle, 1
        ReDim Data(0 To LOF(Handle) - 1) As Byte
        
        Get Handle, , Data
    Close Handle
    
    Call Reader.Wrap(Data)
    
    'map :poop: Header
    MapInfo.MapVersion = Reader.ReadInteger
    Dim i As Long
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
        
            With MapData(X, Y)
                ByFlags = Reader.ReadByte
                    
                .Blocked = ByFlags And 1
                    
                .Graphic(1).GrhIndex = Reader.ReadInteger
                Call InitGrh(.Graphic(1), .Graphic(1).GrhIndex)
                
                For i = 2 To 4
                    If ByFlags And (2 ^ (i - 1)) Then
                        .Graphic(i).GrhIndex = Reader.ReadInteger
                        Call InitGrh(.Graphic(i), .Graphic(i).GrhIndex)
                    Else
                        .Graphic(i).GrhIndex = 0
                    End If
                Next
                
                For i = 4 To 6
                    If (ByFlags And 2 ^ i) Then .Trigger = .Trigger Or 2 ^ (i - 4)
                Next
                
                'Erase NPCs
                If MapData(X, Y).CharIndex > 0 Then
                    Call EraseChar(MapData(X, Y).CharIndex)
                End If
                
                'Erase OBJs
                MapData(X, Y).ObjGrh.GrhIndex = 0
                
            End With
        Next X
    Next Y
    
    MapInfo.Name = ""
    MapInfo.Music = ""
    
    CurMap = Map
End Sub

Function ReadField(ByVal Pos As Integer, ByRef Text As String, ByVal SepASCII As Byte) As String
'*****************************************************************
'Gets a field from a delimited string
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/15/2004
'*****************************************************************
    Dim i As Long
    Dim lastPos As Long
    Dim CurrentPos As Long
    Dim delimiter As String * 1
    
    delimiter = Chr$(SepASCII)
    
    For i = 1 To Pos
        lastPos = CurrentPos
        CurrentPos = InStr(lastPos + 1, Text, delimiter, vbBinaryCompare)
    Next i
    
    If CurrentPos = 0 Then
        ReadField = mid$(Text, lastPos + 1, Len(Text) - lastPos)
    Else
        ReadField = mid$(Text, lastPos + 1, CurrentPos - lastPos - 1)
    End If
End Function

Function FieldCount(ByRef Text As String, ByVal SepASCII As Byte) As Long
'*****************************************************************
'Gets the number of fields in a delimited string
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 07/29/2007
'*****************************************************************
    Dim Count As Long
    Dim curPos As Long
    Dim delimiter As String * 1
    
    If LenB(Text) = 0 Then Exit Function
    
    delimiter = Chr$(SepASCII)
    
    curPos = 0
    
    Do
        curPos = InStr(curPos + 1, Text, delimiter)
        Count = Count + 1
    Loop While curPos <> 0
    
    FieldCount = Count
End Function

Function FileExist(ByVal file As String, ByVal FileType As VbFileAttribute) As Boolean
    FileExist = (Dir$(file, FileType) <> "")
End Function

Sub WriteClientVer()
    Dim hFile As Integer
        
    hFile = FreeFile()
    Open App.path & "\init\Ver.bin" For Binary Access Write Lock Read As #hFile
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    Put #hFile, , CLng(777)
    
    Put #hFile, , CInt(App.Major)
    Put #hFile, , CInt(App.Minor)
    Put #hFile, , CInt(App.Revision)
    
    Close #hFile
End Sub

Public Function IsIp(ByVal Ip As String) As Boolean
    Dim i As Long
    
    For i = 1 To UBound(ServersLst)
        If ServersLst(i).Ip = Ip Then
            IsIp = True
            Exit Function
        End If
    Next i
End Function

Public Sub CargarServidores()
'********************************
'Author: Unknown
'Last Modification: 07/26/07
'Last Modified by: Rapsodius
'Added Instruction "CloseClient" before End so the mutex is cleared
'********************************
On Error GoTo errorH
    Dim f As String
    Dim c As Integer
    Dim i As Long
    
    f = App.path & "\init\sinfo.dat"
    c = Val(GetVar(f, "INIT", "Cant"))
    
    ReDim ServersLst(1 To c) As tServerInfo
    For i = 1 To c
        ServersLst(i).Desc = GetVar(f, "S" & i, "Desc")
        ServersLst(i).Ip = Trim$(GetVar(f, "S" & i, "Ip"))
        ServersLst(i).PassRecPort = CInt(GetVar(f, "S" & i, "P2"))
        ServersLst(i).Puerto = CInt(GetVar(f, "S" & i, "PJ"))
    Next i
    CurServer = 1
Exit Sub

errorH:
    Call MsgBox("Error cargando los servidores, actualicelos de la web", vbCritical + vbOKOnly, "Argentum Online")
    
    Call CloseClient
End Sub

Public Sub InitServersList()
On Error Resume Next
    Dim NumServers As Integer
    Dim i As Integer
    Dim Cont As Integer
    
    i = 1
    
    Do While (ReadField(i, RawServersList, Asc(";")) <> "")
        i = i + 1
        Cont = Cont + 1
    Loop
    
    ReDim ServersLst(1 To Cont) As tServerInfo
    
    For i = 1 To Cont
        Dim cur$
        cur$ = ReadField(i, RawServersList, Asc(";"))
        ServersLst(i).Ip = ReadField(1, cur$, Asc(":"))
        ServersLst(i).Puerto = ReadField(2, cur$, Asc(":"))
        ServersLst(i).Desc = ReadField(4, cur$, Asc(":"))
        ServersLst(i).PassRecPort = ReadField(3, cur$, Asc(":"))
    Next i
    
    CurServer = 1
End Sub

Sub Main()
    Call WriteClientVer
    Call InitColours
        
    'Load config file
    If FileExist(App.path & "\init\Inicio.con", vbNormal) Then
        Config_Inicio = LeerGameIni()
    End If
    
    CurServerIP = "127.0.0.1"
    CurServerPort = 7666
    
    'Load ao.dat config file
    Call LoadClientSetup

    
    'Read command line. Do it AFTER config file is loaded to prevent this from
    'canceling the effects of "/nores" option.
    Call LeerLineaComandos
    
    'usaremos esto para ayudar en los parches
    Call SaveSetting("ArgentumOnlineCliente", "Init", "Path", App.path & "\")
    
    ChDrive App.path
    ChDir App.path

    MD5HushYo = "0123456789abcdef"  'We aren't using a real MD5
    
    tipf = Config_Inicio.tip
    
    'Set resolution BEFORE the loading form is displayed, therefore it will be centered.
    'Call Resolution.SetResolution
    
    ' Mouse Pointer (Loaded before opening any form with buttons in it)
    If FileExist(DirExtras & "Hand.ico", vbArchive) Then _
        Set picMouseIcon = LoadPicture(DirExtras & "Hand.ico")
    
    frmCargando.Show
    frmCargando.Refresh
    
    'frmConnect.version = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
    Call AddtoRichTextBox(frmCargando.Status, "Buscando servidores... ", 255, 255, 255, True, False, True)

    Call CargarServidores
'TODO : esto de ServerRecibidos no se podr�a sacar???
    ServersRecibidos = True
    
    Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, False)
    Call AddtoRichTextBox(frmCargando.Status, "Iniciando constantes... ", 255, 255, 255, True, False, True)
    
    Call InicializarNombres
    
    ' Initialize FONTTYPES
    Call Protocol.InitFonts
    
    With frmConnect
        '.txtNombre = Config_Inicio.Name
        '.txtNombre.SelStart = 0
        '.txtNombre.SelLength = Len(.txtNombre)
    End With
    
    Call EstablecerRecompensas
    Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.Status, "Iniciando motor gr�fico... ", 255, 255, 255, True, False, True)
    
    prgRun = True
    
    If Not InitTileEngine(frmMain.hWnd, 149, 13, 32, 32, 17, 23, 7, 8, 8, 0.018) Then
        Call CloseClient
    End If
    
    'Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.Status, "Creando animaciones extra... ", 255, 255, 255, True, False, True)
    
    Call CargarTips
    
UserMap = 1
    
    Call CargarArrayLluvia
    Call CargarAnimArmas
    Call CargarAnimEscudos
    Call CargarColores
    
    Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.Status, "Iniciando DirectSound... ", 255, 255, 255, True, False, True)
    
    'Inicializamos el sonido
    'call 'audio.Initialize(DirectX, frmMain.hwnd, App.path & "\" & Config_Inicio.DirSonidos & "\", App.path & "\" & Config_Inicio.DirMusica & "\")
    'Enable / Disable audio
    'audio.MusicActivated = Not ClientSetup.bNoMusic
    'audio.SoundActivated = Not ClientSetup.bNoSound
    'audio.SoundEffectsActivated = Not ClientSetup.bNoSoundEffects
    
    'Inicializamos el inventario gr�fico
    Call Inventario.Initialize(800, 223, 160, 160, MAX_INVENTORY_SLOTS)
    
    'call 'audio.MusicMP3Play(App.path & "\MP3\" & MP3_Inicio & ".mp3")
    
    Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.Status, "                    �Bienvenido a Argentum Online!", 255, 255, 255, True, False, True)
    
    'Give the user enough time to read the welcome text
    Call Sleep(500)
    
    Unload frmCargando
        
    frmMain.Socket1.Startup

    frmConnect.Visible = True
    
    'Inicializaci�n de variables globales
    PrimeraVez = True
    pausa = False
    
    'Set the intervals of timers
    Call MainTimer.SetInterval(TimersIndex.Attack, INT_ATTACK)
    Call MainTimer.SetInterval(TimersIndex.Work, INT_WORK)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithU, INT_USEITEMU)
    Call MainTimer.SetInterval(TimersIndex.UseItemWithDblClick, INT_USEITEMDCK)
    Call MainTimer.SetInterval(TimersIndex.SendRPU, INT_SENTRPU)
    Call MainTimer.SetInterval(TimersIndex.CastSpell, INT_CAST_SPELL)
    Call MainTimer.SetInterval(TimersIndex.Arrows, INT_ARROWS)
    Call MainTimer.SetInterval(TimersIndex.CastAttack, INT_CAST_ATTACK)
    
    frmMain.macrotrabajo.Interval = INT_MACRO_TRABAJO
    frmMain.macrotrabajo.Enabled = False
    
   'Init timers
    Call MainTimer.Start(TimersIndex.Attack)
    Call MainTimer.Start(TimersIndex.Work)
    Call MainTimer.Start(TimersIndex.UseItemWithU)
    Call MainTimer.Start(TimersIndex.UseItemWithDblClick)
    Call MainTimer.Start(TimersIndex.SendRPU)
    Call MainTimer.Start(TimersIndex.CastSpell)
    Call MainTimer.Start(TimersIndex.Arrows)
    Call MainTimer.Start(TimersIndex.CastAttack)
    
    'Set the dialog's font
    Dialogos.Font = frmMain.Font
    'DialogosClanes.font = frmMain.font
    
    lFrameTimer = GetTickCount
    
    ' Load the form for screenshots
    Call Load(frmScreenshots)
        
    Do While prgRun
        'S�lo dibujamos si la ventana no est� minimizada
        If frmMain.WindowState <> 1 And frmMain.Visible Then
            Call ShowNextFrame(frmMain.Top, frmMain.Left, frmMain.MouseX, frmMain.MouseY)
            
            'Play ambient sounds
            Call RenderSounds
            
            Call CheckKeys
        End If
        'FPS Counter - mostramos las FPS
        'If GetTickCount - lFrameTimer >= 1000 Then
        '    If FPSFLAG Then frmMain.lblFPS.Caption = Mod_TileEngine.FPS
            
        '    lFrameTimer = GetTickCount
        'End If
        
       ' If GetTickCount - Count = 1000 Then
        '        Call SendData(SendTarget.toMap, UserIndex, PrepareMessageCountdown(Count))
        '        GetTickCount = Count
        '    End If
        ' If there is anything to be sent, we send it
        Call FlushBuffer
        
        DoEvents
    Loop
    
    Call CloseClient
End Sub

Sub WriteVar(ByVal file As String, ByVal Main As String, ByVal Var As String, ByVal Value As String)
'*****************************************************************
'Writes a var to a text file
'*****************************************************************
    writeprivateprofilestring Main, Var, Value, file
End Sub

Function GetVar(ByVal file As String, ByVal Main As String, ByVal Var As String) As String
'*****************************************************************
'Gets a Var from a text file
'*****************************************************************
    Dim sSpaces As String ' This will hold the input that the program will retrieve
    
    sSpaces = Space$(500) ' This tells the computer how long the longest string can be. If you want, you can change the number 100 to any number you wish
    
    getprivateprofilestring Main, Var, vbNullString, sSpaces, Len(sSpaces), file
    
    GetVar = RTrim$(sSpaces)
    GetVar = Left$(GetVar, Len(GetVar) - 1)
End Function

'[CODE 002]:MatuX
'
'  Funci�n para chequear el email
'
'  Corregida por Maraxus para que reconozca como v�lidas casillas con puntos antes de la arroba y evitar un chequeo innecesario
Public Function CheckMailString(ByVal sString As String) As Boolean
On Error GoTo errHnd
    Dim lPos  As Long
    Dim lX    As Long
    Dim iAsc  As Integer
    
    '1er test: Busca un simbolo @
    lPos = InStr(sString, "@")
    If (lPos <> 0) Then
        '2do test: Busca un simbolo . despu�s de @ + 1
        If Not (InStr(lPos, sString, ".", vbBinaryCompare) > lPos + 1) Then _
            Exit Function
        
        '3er test: Recorre todos los caracteres y los val�da
        For lX = 0 To Len(sString) - 1
            If Not (lX = (lPos - 1)) Then   'No chequeamos la '@'
                iAsc = Asc(mid$(sString, (lX + 1), 1))
                If Not CMSValidateChar_(iAsc) Then _
                    Exit Function
            End If
        Next lX
        
        'Finale
        CheckMailString = True
    End If
errHnd:
End Function

'  Corregida por Maraxus para que reconozca como v�lidas casillas con puntos antes de la arroba
Private Function CMSValidateChar_(ByVal iAsc As Integer) As Boolean
    CMSValidateChar_ = (iAsc >= 48 And iAsc <= 57) Or _
                        (iAsc >= 65 And iAsc <= 90) Or _
                        (iAsc >= 97 And iAsc <= 122) Or _
                        (iAsc = 95) Or (iAsc = 45) Or (iAsc = 46)
End Function

'TODO : como todo lo relativo a mapas, no tiene nada que hacer ac�....
Function HayAgua(ByVal X As Integer, ByVal Y As Integer) As Boolean
    HayAgua = ((MapData(X, Y).Graphic(1).GrhIndex >= 1505 And MapData(X, Y).Graphic(1).GrhIndex <= 1520) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 5665 And MapData(X, Y).Graphic(1).GrhIndex <= 5680) Or _
            (MapData(X, Y).Graphic(1).GrhIndex >= 13547 And MapData(X, Y).Graphic(1).GrhIndex <= 13562)) And _
                MapData(X, Y).Graphic(2).GrhIndex = 0
                
End Function

Public Sub ShowSendTxt()
    If Not frmCantidad.Visible Then
        'todo
        'frmMain.SendTxt.Visible = True
        'frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub ShowSendCMSGTxt()
    If Not frmCantidad.Visible Then
    'todo
        'frmMain.SendCMSTXT.Visible = True
        'frmMain.SendCMSTXT.SetFocus
    End If
End Sub

''
' Checks the command line parameters, if you are running Ao with /nores command and checks the AoUpdate parameters
'
'

Public Sub LeerLineaComandos()
'*************************************************
'Author: Unknown
'Last modified: 25/11/2008 (BrianPr)
'
'*************************************************
    Dim T() As String
    Dim i As Long
    
    Dim UpToDate As Boolean

    'Parseo los comandos
    T = Split(Command, " ")
    For i = LBound(T) To UBound(T)
        Select Case UCase$(T(i))
            Case "/NORES" 'no cambiar la resolucion
                NoRes = True
            Case "/UPTODATE"
                UpToDate = True
        End Select
    Next i
    
    'Call AoUpdate(UpToDate, NoRes) ' www.gs-zone.org
End Sub

Private Sub LoadClientSetup()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/19/09
'11/19/09: Pato - Is optional show the frmGuildNews form
'**************************************************************
    Dim fHandle As Integer
    
    If FileExist(App.path & "\init\ao.dat", vbArchive) Then
        fHandle = FreeFile
        
        Open App.path & "\init\ao.dat" For Binary Access Read Lock Write As fHandle
            Get fHandle, , ClientSetup
        Close fHandle
    Else
        'Use dynamic by default
        ClientSetup.bDinamic = True
    End If
    
    NoRes = ClientSetup.bNoRes
    
    If InStr(1, ClientSetup.sGraficos, "Graficos") Then
        GraphicsFile = ClientSetup.sGraficos
    Else
        GraphicsFile = "Graficos3.ind"
    End If
    
 '   ClientSetup.bGuildNews = Not ClientSetup.bGuildNews
  '  DialogosClanes.Activo = Not ClientSetup.bGldMsgConsole
  '  DialogosClanes.CantidadDialogos = ClientSetup.bCantMsgs
End Sub

Private Sub SaveClientSetup()
'**************************************************************
'Author: Torres Patricio (Pato)
'Last Modify Date: 03/11/10
'
'**************************************************************
    Dim fHandle As Integer
    
    fHandle = FreeFile
    
    'ClientSetup.bNoMusic = Not 'audio.MusicActivated
    'ClientSetup.bNoSound = Not 'audio.SoundActivated
    'ClientSetup.bNoSoundEffects = Not 'audio.SoundEffectsActivated
   ' ClientSetup.bGuildNews = Not ClientSetup.bGuildNews
  '  ClientSetup.bGldMsgConsole = Not DialogosClanes.Activo
    'ClientSetup.bCantMsgs = DialogosClanes.CantidadDialogos
    
    Open App.path & "\init\ao.dat" For Binary As fHandle
        Put fHandle, , ClientSetup
    Close fHandle
End Sub

Private Sub InicializarNombres()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Inicializa los nombres de razas, ciudades, clases, skills, atributos, etc.
'**************************************************************
    Ciudades(eCiudad.cUllathorpe) = "Ullathorpe"
    Ciudades(eCiudad.cNix) = "Nix"
    Ciudades(eCiudad.cBanderbill) = "Banderbill"
    Ciudades(eCiudad.cLindos) = "Lindos"
    Ciudades(eCiudad.cArghal) = "Argh�l"
    
    ListaRazas(eRaza.Humano) = "Humano"
    ListaRazas(eRaza.Elfo) = "Elfo"
    ListaRazas(eRaza.ElfoOscuro) = "Elfo Oscuro"
    ListaRazas(eRaza.Gnomo) = "Gnomo"
    ListaRazas(eRaza.Enano) = "Enano"

    ListaClases(eClass.Ciudadano) = "Ciudadano"
    ListaClases(eClass.Trabajador) = "Trabajador"
    ListaClases(eClass.Experto_Minerales) = "Experto en minerales"
    ListaClases(eClass.MINERO) = "Minero"
    ListaClases(eClass.HERRERO) = "Herrero"
    ListaClases(eClass.Experto_Madera) = "Experto en uso de madera"
    ListaClases(eClass.TALADOR) = "Le�ador"
    ListaClases(eClass.CARPINTERO) = "Carpintero"
    ListaClases(eClass.PESCADOR) = "Pescador"
    ListaClases(eClass.Sastre) = "Sastre"
    ListaClases(eClass.Alquimista) = "Alquimista"
    ListaClases(eClass.Luchador) = "Luchador"
    ListaClases(eClass.Con_Mana) = "Con uso de mana"
    ListaClases(eClass.Hechicero) = "Hechicero"
    ListaClases(eClass.MAGO) = "Mago"
    ListaClases(eClass.NIGROMANTE) = "Nigromante"
    ListaClases(eClass.Orden_Sagrada) = "Orden sagrada"
    ListaClases(eClass.PALADIN) = "Paladin"
    ListaClases(eClass.CLERIGO) = "Clerigo"
    ListaClases(eClass.Naturalista) = "Naturalista"
    ListaClases(eClass.BARDO) = "Bardo"
    ListaClases(eClass.DRUIDA) = "Druida"
    ListaClases(eClass.Sigiloso) = "Sigiloso"
    ListaClases(eClass.ASESINO) = "Asesino"
    ListaClases(eClass.CAZADOR) = "Cazador"
    ListaClases(eClass.Sin_Mana) = "Sin uso de mana"
    ListaClases(eClass.ARQUERO) = "Arquero"
    ListaClases(eClass.GUERRERO) = "Guerrero"
    ListaClases(eClass.Caballero) = "Caballero"
    ListaClases(eClass.Bandido) = "Bandido"
    ListaClases(eClass.PIRATA) = "Pirata"
    ListaClases(eClass.LADRON) = "Ladron"
    
    SkillsNames(eSkill.Magia) = "Magia"
    SkillsNames(eSkill.Robar) = "Robar"
    SkillsNames(eSkill.Tacticas) = "Evasi�n en combate"
    SkillsNames(eSkill.Armas) = "Combate cuerpo a cuerpo"
    SkillsNames(eSkill.Meditar) = "Meditar"
    SkillsNames(eSkill.Apu�alar) = "Apu�alar"
    SkillsNames(eSkill.Ocultarse) = "Ocultarse"
    SkillsNames(eSkill.Supervivencia) = "Supervivencia"
    SkillsNames(eSkill.Talar) = "Talar �rboles"
    SkillsNames(eSkill.Comerciar) = "Comercio"
    SkillsNames(eSkill.Defensa) = "Defensa con escudos"
    SkillsNames(eSkill.Pesca) = "Pesca"
    SkillsNames(eSkill.Mineria) = "Mineria"
    SkillsNames(eSkill.Carpinteria) = "Carpinteria"
    SkillsNames(eSkill.Herreria) = "Herreria"
    SkillsNames(eSkill.Liderazgo) = "Liderazgo"
    SkillsNames(eSkill.Domar) = "Domar animales"
    SkillsNames(eSkill.Proyectiles) = "Combate a distancia"
    SkillsNames(eSkill.Wrestling) = "Combate sin armas"
    SkillsNames(eSkill.Navegacion) = "Navegacion"
    SkillsNames(eSkill.Sastreria) = "Sastrer�a"
    SkillsNames(eSkill.Resis) = "Resistencia M�gica"
    
    AtributosNames(eAtributos.Fuerza) = "Fuerza"
    AtributosNames(eAtributos.Agilidad) = "Agilidad"
    AtributosNames(eAtributos.Inteligencia) = "Inteligencia"
    AtributosNames(eAtributos.Carisma) = "Carisma"
    AtributosNames(eAtributos.Constitucion) = "Constitucion"
End Sub

''
' Removes all text from the console and dialogs

Public Sub CleanDialogs()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 11/27/2005
'Removes all text from the console and dialogs
'**************************************************************
    'Clean console and dialogs
    'todo
    'frmMain.RecTxt.Text = vbNullString
    
  '  Call DialogosClanes.RemoveDialogs
    
    Call Dialogos.RemoveAllDialogs
End Sub

Public Sub CloseClient()
'**************************************************************
'Author: Juan Mart�n Sotuyo Dodero (Maraxus)
'Last Modify Date: 8/14/2007
'Frees all used resources, cleans up and leaves
'**************************************************************
    ' Allow new instances of the client to be opened
    Call PrevInstance.ReleaseInstance
    
    EngineRun = False
    frmCargando.Show
    Call AddtoRichTextBox(frmCargando.Status, "Liberando recursos...", 0, 0, 0, 0, 0, 0)
    
    'Stop tile engine
    Call DeinitTileEngine
    
    Call SaveClientSetup
    
    'Destruimos los objetos p�blicos creados
    Set CustomMessages = Nothing
    Set CustomKeys = Nothing
    'Set SurfaceDB = Nothing
    Set Dialogos = Nothing
  '  Set DialogosClanes = Nothing
    'Set Audio = Nothing
    Set Inventario = Nothing
    Set MainTimer = Nothing
    Set incomingData = Nothing
    Set outgoingData = Nothing
    
    Call UnloadAllForms
    
    'Actualizar tip
    Config_Inicio.tip = tipf
    Call EscribirGameIni(Config_Inicio)
    End
End Sub

Public Function esGM(CharIndex As Integer) As Boolean
esGM = False
If charlist(CharIndex).priv >= 1 And charlist(CharIndex).priv <= 5 Or charlist(CharIndex).priv = 25 Then _
    esGM = True

End Function

Public Function getTagPosition(ByVal Nick As String) As Integer
Dim buf As Integer
buf = InStr(Nick, "<")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
buf = InStr(Nick, "[")
If buf > 0 Then
    getTagPosition = buf
    Exit Function
End If
getTagPosition = Len(Nick) + 2
End Function

Public Function getStrenghtColor() As Long
Dim M As Long
M = 255 / MAXATRIBUTOS
getStrenghtColor = RGB(255 - (M * UserFuerza), (M * UserFuerza), 0)
End Function
Public Function getDexterityColor() As Long
Dim M As Long
M = 255 / MAXATRIBUTOS
getDexterityColor = RGB(255, M * UserAgilidad, 0)
End Function

Public Function getCharIndexByName(ByVal Name As String) As Integer
Dim i As Long
For i = 1 To LastChar
    If charlist(i).Nombre = Name Then
        getCharIndexByName = i
        Exit Function
    End If
Next i
End Function

Public Function EsAnuncio(ByVal ForumType As Byte) As Boolean
'***************************************************
'Author: ZaMa
'Last Modification: 22/02/2010
'Returns true if the post is sticky.
'***************************************************
    Select Case ForumType
        Case eForumMsgType.ieCAOS_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieGENERAL_STICKY
            EsAnuncio = True
            
        Case eForumMsgType.ieREAL_STICKY
            EsAnuncio = True
            
    End Select
    
End Function

Public Function ForumAlignment(ByVal yForumType As Byte) As Byte
'***************************************************
'Author: ZaMa
'Last Modification: 01/03/2010
'Returns the forum alignment.
'***************************************************
    Select Case yForumType
        Case eForumMsgType.ieCAOS, eForumMsgType.ieCAOS_STICKY
            ForumAlignment = eForumType.ieCAOS
            
        Case eForumMsgType.ieGeneral, eForumMsgType.ieGENERAL_STICKY
            ForumAlignment = eForumType.ieGeneral
            
        Case eForumMsgType.ieREAL, eForumMsgType.ieREAL_STICKY
            ForumAlignment = eForumType.ieREAL
            
    End Select
    
End Function

Public Sub LogError(Desc As String)
'***************************************************
'Author: Unknown
'Last Modification: -
'
'***************************************************

On Error GoTo ErrHandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.path & "\errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Desc
    Close #nfile
    
    Exit Sub

ErrHandler:

End Sub

Public Function ReadFile(FileName As String, Optional Size As Long = -1) As Byte()

    Dim wFile As Integer

    wFile = FreeFile
    Open FileName For Binary Access Read As wFile

    If LOF(wFile) > 0 Then
        
        Size = LOF(wFile)
        ReDim ReadFile(0 To LOF(wFile) - 1)
        Get wFile, , ReadFile

    End If

    Close #wFile

End Function

Public Sub EstablecerRecompensas()

ReDim Recompensas(1 To NUMCLASES, 1 To 3, 1 To 2) As tRecompensa

Recompensas(eClass.MINERO, 1, 1).Name = "Fortaleza del Trabajador"
Recompensas(eClass.MINERO, 1, 1).Descripcion = "Aumenta la vida en 120 puntos."

Recompensas(eClass.MINERO, 1, 2).Name = "Suerte de Novato"
Recompensas(eClass.MINERO, 1, 2).Descripcion = "Al morir hay 20% de probabilidad de no perder los minerales."

Recompensas(eClass.MINERO, 2, 1).Name = "Destrucci�n M�gica"
Recompensas(eClass.MINERO, 2, 1).Descripcion = "Inmunidad al paralisis lanzado por otros usuarios."

Recompensas(eClass.MINERO, 2, 2).Name = "Pica Fuerte"
Recompensas(eClass.MINERO, 2, 2).Descripcion = "Permite minar 20% m�s cantidad de hierro y la plata."

Recompensas(eClass.MINERO, 3, 1).Name = "Gremio del Trabajador"
Recompensas(eClass.MINERO, 3, 1).Descripcion = "Permite minar 20% m�s cantidad de oro."

Recompensas(eClass.MINERO, 3, 2).Name = "Pico de la Suerte"
Recompensas(eClass.MINERO, 3, 2).Descripcion = "Al morir hay 30% de probabilidad de que no perder los minerales (acumulativo con Suerte de Novato.)"


Recompensas(eClass.HERRERO, 1, 1).Name = "Yunque Rojizo"
Recompensas(eClass.HERRERO, 1, 1).Descripcion = "25% de probabilidad de gastar la mitad de lingotes en la creaci�n de objetos (Solo aplicable a armas y armaduras)."

Recompensas(eClass.HERRERO, 1, 2).Name = "Maestro de la Forja"
Recompensas(eClass.HERRERO, 1, 2).Descripcion = "Reduce los costos de cascos y escudos a un 50%."

Recompensas(eClass.HERRERO, 2, 1).Name = "Experto en Filos"
Recompensas(eClass.HERRERO, 2, 1).Descripcion = "Permite crear las mejores armas (Espada Neithan, Espada Neithan + 1, Espada de Plata + 1 y Daga Infernal)."

Recompensas(eClass.HERRERO, 2, 2).Name = "Experto en Corazas"
Recompensas(eClass.HERRERO, 2, 2).Descripcion = "Permite crear las mejores armaduras (Armaduras de las Tinieblas, Armadura Legendaria y Armaduras del Drag�n)."

Recompensas(eClass.HERRERO, 3, 1).Name = "Fundir Metal"
Recompensas(eClass.HERRERO, 3, 1).Descripcion = "Reduce a un 50% la cantidad de lingotes utilizados en fabricaci�n de Armas y Armaduras (acumulable con Yunque Rojizo)."

Recompensas(eClass.HERRERO, 3, 2).Name = "Trabajo en Serie"
Recompensas(eClass.HERRERO, 3, 2).Descripcion = "10% de probabilidad de crear el doble de objetos de los asignados con la misma cantidad de lingotes."


Recompensas(eClass.TALADOR, 1, 1).Name = "M�sculos Fornidos"
Recompensas(eClass.TALADOR, 1, 1).Descripcion = "Permite talar 20% m�s cantidad de madera."

Recompensas(eClass.TALADOR, 1, 2).Name = "Tiempos de Calma"
Recompensas(eClass.TALADOR, 1, 2).Descripcion = "Evita tener hambre y sed."


Recompensas(eClass.CARPINTERO, 1, 1).Name = "Experto en Arcos"
Recompensas(eClass.CARPINTERO, 1, 1).Descripcion = "Permite la creaci�n de los mejores arcos (�lfico y de las Tinieblas)."

Recompensas(eClass.CARPINTERO, 1, 2).Name = "Experto de Varas"
Recompensas(eClass.CARPINTERO, 1, 2).Descripcion = "Permite la creaci�n de las mejores varas (Engarzadas)."

Recompensas(eClass.CARPINTERO, 2, 1).Name = "Fila de Le�a"
Recompensas(eClass.CARPINTERO, 2, 1).Descripcion = "Aumenta la creaci�n de flechas a 20 por vez."

Recompensas(eClass.CARPINTERO, 2, 2).Name = "Esp�ritu de Navegante"
Recompensas(eClass.CARPINTERO, 2, 2).Descripcion = "Reduce en un 20% el coste de madera de las barcas."


Recompensas(eClass.PESCADOR, 1, 1).Name = "Favor de los Dioses"
Recompensas(eClass.PESCADOR, 1, 1).Descripcion = "Pescar 20% m�s cantidad de pescados."

Recompensas(eClass.PESCADOR, 1, 2).Name = "Pesca en Alta Mar"
Recompensas(eClass.PESCADOR, 1, 2).Descripcion = "Al pescar en barca hay 10% de probabilidad de obtener pescados m�s caros."


Recompensas(eClass.MAGO, 1, 1).Name = "Pociones de Esp�ritu"
Recompensas(eClass.MAGO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(eClass.MAGO, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.MAGO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.MAGO, 2, 1).Name = "Vitalidad"
Recompensas(eClass.MAGO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(eClass.MAGO, 2, 2).Name = "Fortaleza Mental"
Recompensas(eClass.MAGO, 2, 2).Descripcion = "Libera el limite de mana m�ximo."

Recompensas(eClass.MAGO, 3, 1).Name = "Furia del Rel�mpago"
Recompensas(eClass.MAGO, 3, 1).Descripcion = "Aumenta el da�o base m�ximo de la Descarga El�ctrica en 10 puntos."

Recompensas(eClass.MAGO, 3, 2).Name = "Destrucci�n"
Recompensas(eClass.MAGO, 3, 2).Descripcion = "Aumenta el da�o base m�nimo del Apocalipsis en 10 puntos."

Recompensas(eClass.NIGROMANTE, 1, 1).Name = "Pociones de Esp�ritu"
Recompensas(eClass.NIGROMANTE, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(eClass.NIGROMANTE, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.NIGROMANTE, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.NIGROMANTE, 2, 1).Name = "Vida del Invocador"
Recompensas(eClass.NIGROMANTE, 2, 1).Descripcion = "Aumenta la vida en 15 puntos."

Recompensas(eClass.NIGROMANTE, 2, 2).Name = "Alma del Invocador"
Recompensas(eClass.NIGROMANTE, 2, 2).Descripcion = "Aumenta el mana en 40 puntos."

Recompensas(eClass.NIGROMANTE, 3, 1).Name = "Semillas de las Almas"
Recompensas(eClass.NIGROMANTE, 3, 1).Descripcion = "Aumenta el da�o base m�nimo de la magia en 10 puntos."

Recompensas(eClass.NIGROMANTE, 3, 2).Name = "Bloqueo de las Almas"
Recompensas(eClass.NIGROMANTE, 3, 2).Descripcion = "Aumenta la evasi�n en un 5%."


Recompensas(eClass.PALADIN, 1, 1).Name = "Pociones de Esp�ritu"
Recompensas(eClass.PALADIN, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(eClass.PALADIN, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.PALADIN, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.PALADIN, 2, 1).Name = "Aura de Vitalidad"
Recompensas(eClass.PALADIN, 2, 1).Descripcion = "Aumenta la vida en 5 puntos y el mana en 10 puntos."

Recompensas(eClass.PALADIN, 2, 2).Name = "Aura de Esp�ritu"
Recompensas(eClass.PALADIN, 2, 2).Descripcion = "Aumenta el mana en 30 puntos."

Recompensas(eClass.PALADIN, 3, 1).Name = "Gracia Divina"
Recompensas(eClass.PALADIN, 3, 1).Descripcion = "Reduce el coste de mana de Remover Paralisis a 250 puntos."

Recompensas(eClass.PALADIN, 3, 2).Name = "Favor de los Enanos"
Recompensas(eClass.PALADIN, 3, 2).Descripcion = "Aumenta en 5% la posibilidad de golpear al enemigo con armas cuerpo a cuerpo."

Recompensas(eClass.CLERIGO, 1, 1).Name = "Pociones de Esp�ritu"
Recompensas(eClass.CLERIGO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(eClass.CLERIGO, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.CLERIGO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.CLERIGO, 2, 1).Name = "Signo Vital"
Recompensas(eClass.CLERIGO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(eClass.CLERIGO, 2, 2).Name = "Esp�ritu de Sacerdote"
Recompensas(eClass.CLERIGO, 2, 2).Descripcion = "Aumenta el mana en 50 puntos."

Recompensas(eClass.CLERIGO, 3, 1).Name = "Sacerdote Experto"
Recompensas(eClass.CLERIGO, 3, 1).Descripcion = "Aumenta la cura base de Curar Heridas Graves en 20 puntos."

Recompensas(eClass.CLERIGO, 3, 2).Name = "Alzamientos de Almas"
Recompensas(eClass.CLERIGO, 3, 2).Descripcion = "El hechizo de Resucitar cura a las personas con su mana, energ�a, hambre y sed llenas y cuesta 1.100 de mana."

Recompensas(eClass.BARDO, 1, 1).Name = "Pociones de Esp�ritu"
Recompensas(eClass.BARDO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(eClass.BARDO, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.BARDO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.BARDO, 2, 1).Name = "Melod�a Vital"
Recompensas(eClass.BARDO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(eClass.BARDO, 2, 2).Name = "Melod�a de la Meditaci�n"
Recompensas(eClass.BARDO, 2, 2).Descripcion = "Aumenta el mana en 50 puntos."

Recompensas(eClass.BARDO, 3, 1).Name = "Concentraci�n"
Recompensas(eClass.BARDO, 3, 1).Descripcion = "Aumenta la probabilidad de Apu�alar a un 20% (con 100 skill)."

Recompensas(eClass.BARDO, 3, 2).Name = "Melod�a Ca�tica"
Recompensas(eClass.BARDO, 3, 2).Descripcion = "Aumenta el da�o base del Apocalipsis y la Descarga Electrica en 5 puntos."


Recompensas(eClass.DRUIDA, 1, 1).Name = "Pociones de Esp�ritu"
Recompensas(eClass.DRUIDA, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(eClass.DRUIDA, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.DRUIDA, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.DRUIDA, 2, 1).Name = "Grifo de la Vida"
Recompensas(eClass.DRUIDA, 2, 1).Descripcion = "Aumenta la vida en 15 puntos."

Recompensas(eClass.DRUIDA, 2, 2).Name = "Poder del Alma"
Recompensas(eClass.DRUIDA, 2, 2).Descripcion = "Aumenta el mana en 40 puntos."

Recompensas(eClass.DRUIDA, 3, 1).Name = "Ra�ces de la Naturaleza"
Recompensas(eClass.DRUIDA, 3, 1).Descripcion = "Reduce el coste de mana de Inmovilizar a 250 puntos."

Recompensas(eClass.DRUIDA, 3, 2).Name = "Fortaleza Natural"
Recompensas(eClass.DRUIDA, 3, 2).Descripcion = "Aumenta la vida de los elementales invocados en 75 puntos."


Recompensas(eClass.ASESINO, 1, 1).Name = "Pociones de Esp�ritu"
Recompensas(eClass.ASESINO, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(eClass.ASESINO, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.ASESINO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.ASESINO, 2, 1).Name = "Sombra de Vida"
Recompensas(eClass.ASESINO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(eClass.ASESINO, 2, 2).Name = "Sombra M�gica"
Recompensas(eClass.ASESINO, 2, 2).Descripcion = "Aumenta el mana en 30 puntos."

Recompensas(eClass.ASESINO, 3, 1).Name = "Daga Mortal"
Recompensas(eClass.ASESINO, 3, 1).Descripcion = "Aumenta el da�o de Apu�alar a un 70% m�s que el golpe."

Recompensas(eClass.ASESINO, 3, 2).Name = "Punteria mortal"
Recompensas(eClass.ASESINO, 3, 2).Descripcion = "Las chances de apu�alar suben a 25% (Con 100 skills)."


Recompensas(eClass.CAZADOR, 1, 1).Name = "Pociones de Esp�ritu"
Recompensas(eClass.CAZADOR, 1, 1).Descripcion = "1.000 pociones azules que no caen al morir."

Recompensas(eClass.CAZADOR, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.CAZADOR, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.CAZADOR, 2, 1).Name = "Fortaleza del Oso"
Recompensas(eClass.CAZADOR, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(eClass.CAZADOR, 2, 2).Name = "Fortaleza del Leviat�n"
Recompensas(eClass.CAZADOR, 2, 2).Descripcion = "Aumenta el mana en 50 puntos."

Recompensas(eClass.CAZADOR, 3, 1).Name = "Precisi�n"
Recompensas(eClass.CAZADOR, 3, 1).Descripcion = "Aumenta la punter�a con arco en un 10%."

Recompensas(eClass.CAZADOR, 3, 2).Name = "Tiro Preciso"
Recompensas(eClass.CAZADOR, 3, 2).Descripcion = "Las flechas que golpeen la cabeza ignoran la defensa del casco."


Recompensas(eClass.ARQUERO, 1, 1).Name = "Flechas Mortales"
Recompensas(eClass.ARQUERO, 1, 1).Descripcion = "1.500 flechas que caen al morir."

Recompensas(eClass.ARQUERO, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.ARQUERO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.ARQUERO, 2, 1).Name = "Vitalidad �lfica"
Recompensas(eClass.ARQUERO, 2, 1).Descripcion = "Aumenta la vida en 10 puntos."

Recompensas(eClass.ARQUERO, 2, 2).Name = "Paso �lfico"
Recompensas(eClass.ARQUERO, 2, 2).Descripcion = "Aumenta la evasi�n en un 5%."

Recompensas(eClass.ARQUERO, 3, 1).Name = "Ojo del �guila"
Recompensas(eClass.ARQUERO, 3, 1).Descripcion = "Aumenta la punter�a con arco en un 5%."

Recompensas(eClass.ARQUERO, 3, 2).Name = "Disparo �lfico"
Recompensas(eClass.ARQUERO, 3, 2).Descripcion = "Aumenta el da�o base m�nimo de las flechas en 5 puntos y el m�ximo en 3 puntos."


Recompensas(eClass.GUERRERO, 1, 1).Name = "Pociones de Poder"
Recompensas(eClass.GUERRERO, 1, 1).Descripcion = "80 pociones verdes y 100 amarillas que no caen al morir."

Recompensas(eClass.GUERRERO, 1, 2).Name = "Pociones de Vida"
Recompensas(eClass.GUERRERO, 1, 2).Descripcion = "1.000 pociones rojas que no caen al morir."

Recompensas(eClass.GUERRERO, 2, 1).Name = "Vida del Mamut"
Recompensas(eClass.GUERRERO, 2, 1).Descripcion = "Aumenta la vida en 5 puntos."

Recompensas(eClass.GUERRERO, 2, 2).Name = "Piel de Piedra"
Recompensas(eClass.GUERRERO, 2, 2).Descripcion = "Aumenta la defensa permanentemente en 2 puntos."

Recompensas(eClass.GUERRERO, 3, 1).Name = "Cuerda Tensa"
Recompensas(eClass.GUERRERO, 3, 1).Descripcion = "Aumenta la punter�a con arco en un 10%."

Recompensas(eClass.GUERRERO, 3, 2).Name = "Resistencia M�gica"
Recompensas(eClass.GUERRERO, 3, 2).Descripcion = "Reduce la duraci�n de la par�lisis de un minuto a 45 segundos."


Recompensas(eClass.PIRATA, 1, 1).Name = "Marejada Vital"
Recompensas(eClass.PIRATA, 1, 1).Descripcion = "Aumenta la vida en 20 puntos."

Recompensas(eClass.PIRATA, 1, 2).Name = "Aventurero Arriesgado"
Recompensas(eClass.PIRATA, 1, 2).Descripcion = "Permite entrar a los dungeons independientemente del nivel."

Recompensas(eClass.PIRATA, 2, 1).Name = "Riqueza"
Recompensas(eClass.PIRATA, 2, 1).Descripcion = "10% de probabilidad de no perder los objetos al morir."

Recompensas(eClass.PIRATA, 2, 2).Name = "Escamas del Drag�n"
Recompensas(eClass.PIRATA, 2, 2).Descripcion = "Aumenta la vida en 40 puntos."

Recompensas(eClass.PIRATA, 3, 1).Name = "Magia Tab�"
Recompensas(eClass.PIRATA, 3, 1).Descripcion = "Inmunidad a la paralisis."

Recompensas(eClass.PIRATA, 3, 2).Name = "Cuerda de Escape"
Recompensas(eClass.PIRATA, 3, 2).Descripcion = "Permite salir del juego en solo dos segundos."


Recompensas(eClass.LADRON, 1, 1).Name = "Codicia"
Recompensas(eClass.LADRON, 1, 1).Descripcion = "Aumenta en 10% la cantidad de oro robado."

Recompensas(eClass.LADRON, 1, 2).Name = "Manos Sigilosas"
Recompensas(eClass.LADRON, 1, 2).Descripcion = "Aumenta en 5% la probabilidad de robar exitosamente."

Recompensas(eClass.LADRON, 2, 1).Name = "Pies sigilosos"
Recompensas(eClass.LADRON, 2, 1).Descripcion = "Permite moverse mientr�s se est� oculto."

Recompensas(eClass.LADRON, 2, 2).Name = "Ladr�n Experto"
Recompensas(eClass.LADRON, 2, 2).Descripcion = "Permite el robo de objetos (10% de probabilidad)."

Recompensas(eClass.LADRON, 3, 1).Name = "Robo Lejano"
Recompensas(eClass.LADRON, 3, 1).Descripcion = "Permite robar a una distancia de hasta 4 tiles."

Recompensas(eClass.LADRON, 3, 2).Name = "Fundido de Sombra"
Recompensas(eClass.LADRON, 3, 2).Descripcion = "Aumenta en 10% la probabilidad de robar objetos."

End Sub
