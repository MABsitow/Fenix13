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

    Dim loopc As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "armas.dat"
    
    NumWeaponAnims = Val(GetVar(arch, "INIT", "NumArmas"))
    
    ReDim WeaponAnimData(1 To NumWeaponAnims) As WeaponAnimData
    
    For loopc = 1 To NumWeaponAnims
        InitGrh WeaponAnimData(loopc).WeaponWalk(1), Val(GetVar(arch, "ARMA" & loopc, "Dir1")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(2), Val(GetVar(arch, "ARMA" & loopc, "Dir2")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(3), Val(GetVar(arch, "ARMA" & loopc, "Dir3")), 0
        InitGrh WeaponAnimData(loopc).WeaponWalk(4), Val(GetVar(arch, "ARMA" & loopc, "Dir4")), 0
    Next loopc
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
        ColoresPJ(i).R = CByte(GetVar(archivoC, CStr(i), "R"))
        ColoresPJ(i).g = CByte(GetVar(archivoC, CStr(i), "G"))
        ColoresPJ(i).b = CByte(GetVar(archivoC, CStr(i), "B"))
    Next i
    
    ' Crimi
    ColoresPJ(50).R = CByte(GetVar(archivoC, "CR", "R"))
    ColoresPJ(50).g = CByte(GetVar(archivoC, "CR", "G"))
    ColoresPJ(50).b = CByte(GetVar(archivoC, "CR", "B"))
    
    ' Ciuda
    ColoresPJ(49).R = CByte(GetVar(archivoC, "CI", "R"))
    ColoresPJ(49).g = CByte(GetVar(archivoC, "CI", "G"))
    ColoresPJ(49).b = CByte(GetVar(archivoC, "CI", "B"))
    
    ' Neutral
    ColoresPJ(48).R = CByte(GetVar(archivoC, "NE", "R"))
    ColoresPJ(48).g = CByte(GetVar(archivoC, "NE", "G"))
    ColoresPJ(48).b = CByte(GetVar(archivoC, "NE", "B"))
    
    ColoresPJ(47).R = CByte(GetVar(archivoC, "NW", "R"))
    ColoresPJ(47).g = CByte(GetVar(archivoC, "NW", "G"))
    ColoresPJ(47).b = CByte(GetVar(archivoC, "NW", "B"))
    
End Sub

Sub CargarAnimEscudos()
On Error Resume Next

    Dim loopc As Long
    Dim arch As String
    
    arch = App.path & "\init\" & "escudos.dat"
    
    NumEscudosAnims = Val(GetVar(arch, "INIT", "NumEscudos"))
    
    ReDim ShieldAnimData(1 To NumEscudosAnims) As ShieldAnimData
    
    For loopc = 1 To NumEscudosAnims
        InitGrh ShieldAnimData(loopc).ShieldWalk(1), Val(GetVar(arch, "ESC" & loopc, "Dir1")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(2), Val(GetVar(arch, "ESC" & loopc, "Dir2")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(3), Val(GetVar(arch, "ESC" & loopc, "Dir3")), 0
        InitGrh ShieldAnimData(loopc).ShieldWalk(4), Val(GetVar(arch, "ESC" & loopc, "Dir4")), 0
    Next loopc
End Sub

Sub AddtoRichTextBox(ByRef RichTextBox As RichTextBox, ByVal Text As String, Optional ByVal red As Integer = -1, Optional ByVal green As Integer, Optional ByVal blue As Integer, Optional ByVal bold As Boolean = False, Optional ByVal italic As Boolean = False, Optional ByVal bCrLf As Boolean = True)
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
        .SelBold = bold
        .SelItalic = italic
        
        If Not red = -1 Then .SelColor = RGB(red, green, blue)
        
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
    Dim loopc As Long
    
    For loopc = 1 To LastChar
        If charlist(loopc).Active = 1 Then
            MapData(charlist(loopc).Pos.X, charlist(loopc).Pos.Y).CharIndex = loopc
        End If
    Next loopc
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
    Dim loopc As Long
    Dim CharAscii As Integer
    
    If checkemail And UserEmail = "" Then
        MsgBox ("Direcci�n de email invalida")
        Exit Function
    End If
    
    If UserPassword = "" Then
        MsgBox ("Ingrese un password.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserPassword)
        CharAscii = Asc(mid$(UserPassword, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Password inv�lido. El caract�r " & Chr$(CharAscii) & " no est� permitido.")
            Exit Function
        End If
    Next loopc
    
    If UserName = "" Then
        MsgBox ("Ingrese un nombre de personaje.")
        Exit Function
    End If
    
    If Len(UserName) > 30 Then
        MsgBox ("El nombre debe tener menos de 30 letras.")
        Exit Function
    End If
    
    For loopc = 1 To Len(UserName)
        CharAscii = Asc(mid$(UserName, loopc, 1))
        If Not LegalCharacter(CharAscii) Then
            MsgBox ("Nombre inv�lido. El caract�r " & Chr$(CharAscii) & " no est� permitido.")
            Exit Function
        End If
    Next loopc
    
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
    
    frmMain.lblName.Caption = UserName
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
        End If
    Else
        If charlist(UserCharIndex).Heading <> Direccion Then
            Call WriteChangeHeading(Direccion)
        End If
    End If
    
    If frmMain.macrotrabajo.Enabled Then Call frmMain.DesactivarMacroTrabajo
    
    ' Update 3D sounds!
    Call Audio.MoveListener(UserPos.X, UserPos.Y)
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
                frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
                Exit Sub
            End If
            
            'Move Right
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyRight)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(EAST)
                'frmMain.Coord.Caption = "(" & UserMap & "," & UserPos.x & "," & UserPos.y & ")"
                frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
                Exit Sub
            End If
        
            'Move down
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyDown)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(SOUTH)
                frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
                Exit Sub
            End If
        
            'Move left
            If GetKeyState(CustomKeys.BindedKey(eKeyType.mKeyLeft)) < 0 Then
                If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
                Call MoveTo(WEST)
                frmMain.Coord.Caption = UserMap & " X: " & UserPos.X & " Y: " & UserPos.Y
                Exit Sub
            End If
            
            ' We haven't moved - Update 3D sounds!
            Call Audio.MoveListener(UserPos.X, UserPos.Y)
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
                Call Audio.MoveListener(UserPos.X, UserPos.Y)
            End If
            
            If frmMain.TrainingMacro.Enabled Then frmMain.DesactivarMacroHechizos
            'frmMain.Coord.Caption = "(" & UserPos.x & "," & UserPos.y & ")"
            frmMain.Coord.Caption = "X: " & UserPos.X & " Y: " & UserPos.Y
        End If
    End If
End Sub

'CSEH: ErrLog
Sub SwitchMap(ByVal Map As Integer)
    Dim Y As Long
    Dim X As Long
    Dim tempint As Integer
    Dim handle As Integer
    Dim reader As New clsByteBuffer
    Dim data() As Byte
    Dim ByFlags As Byte
        
    handle = FreeFile()
    
    Open DirMapas & "Mapa" & Map & ".mcl" For Binary As handle
        Seek handle, 1
        ReDim data(0 To LOF(handle) - 1) As Byte
        
        Get handle, , data
    Close handle
    
    Call reader.initializeReader(data)
    
    'map :poop: Header
    MapInfo.MapVersion = reader.getInteger
    Dim i As Long
    
    'Load arrays
    For Y = YMinMapSize To YMaxMapSize
        For X = XMinMapSize To XMaxMapSize
        
            With MapData(X, Y)
                ByFlags = reader.getByte
                    
                .Blocked = ByFlags And 1
                    
                .Graphic(1).GrhIndex = reader.getInteger
                Call InitGrh(.Graphic(1), .Graphic(1).GrhIndex)
                
                For i = 2 To 4
                    If ByFlags And (2 ^ (i - 1)) Then
                        .Graphic(i).GrhIndex = reader.getInteger
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
    
    'Load config file
    If FileExist(App.path & "\init\Inicio.con", vbNormal) Then
        Config_Inicio = LeerGameIni()
    End If
    
    CurServerIP = "127.0.0.1"
    CurServerPort = 7666
    
    'Load ao.dat config file
    Call LoadClientSetup
    
    If ClientSetup.bDinamic Then
        Set SurfaceDB = New clsSurfaceManDyn
    Else
        Set SurfaceDB = New clsSurfaceManStatic
    End If
    
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
    
    frmConnect.version = "v" & App.Major & "." & App.Minor & " Build: " & App.Revision
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
        .txtNombre = Config_Inicio.Name
        .txtNombre.SelStart = 0
        .txtNombre.SelLength = Len(.txtNombre)
    End With
    
    Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.Status, "Iniciando motor gr�fico... ", 255, 255, 255, True, False, True)
    
    If Not InitTileEngine(frmMain.hwnd, 149, 13, 32, 32, 13, 17, 9, 8, 8, 0.018) Then
        Call CloseClient
    End If
    
    Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, False)
    
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
    Call Audio.Initialize(DirectX, frmMain.hwnd, App.path & "\" & Config_Inicio.DirSonidos & "\", App.path & "\" & Config_Inicio.DirMusica & "\")
    'Enable / Disable audio
    Audio.MusicActivated = Not ClientSetup.bNoMusic
    Audio.SoundActivated = Not ClientSetup.bNoSound
    Audio.SoundEffectsActivated = Not ClientSetup.bNoSoundEffects
    'Inicializamos el inventario gr�fico
    Call Inventario.Initialize(DirectDraw, frmMain.picInv, MAX_INVENTORY_SLOTS)
    
    Call Audio.MusicMP3Play(App.path & "\MP3\" & MP3_Inicio & ".mp3")
    
    Call AddtoRichTextBox(frmCargando.Status, "Hecho", 255, 0, 0, True, False, False)
    
    Call AddtoRichTextBox(frmCargando.Status, "                    �Bienvenido a Argentum Online!", 255, 255, 255, True, False, True)
    
    'Give the user enough time to read the welcome text
    Call Sleep(500)
    
    Unload frmCargando
        
#If UsarWrench = 1 Then
    frmMain.Socket1.Startup
#End If

    frmConnect.Visible = True
    
    'Inicializaci�n de variables globales
    PrimeraVez = True
    prgRun = True
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
    Dialogos.font = frmMain.font
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
        If GetTickCount - lFrameTimer >= 1000 Then
            If FPSFLAG Then frmMain.lblFPS.Caption = Mod_TileEngine.FPS
            
            lFrameTimer = GetTickCount
        End If
        
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
        frmMain.SendTxt.Visible = True
        frmMain.SendTxt.SetFocus
    End If
End Sub

Public Sub ShowSendCMSGTxt()
    If Not frmCantidad.Visible Then
        frmMain.SendCMSTXT.Visible = True
        frmMain.SendCMSTXT.SetFocus
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
    Dim Patch As String
    
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

''
' Runs AoUpdate if we haven't updated yet, patches aoupdate and runs Client normally if we are updated.
'
' @param UpToDate Specifies if we have checked for updates or not
' @param NoREs Specifies if we have to set nores arg when running the client once again (if the AoUpdate is executed).

Private Sub AoUpdate(ByVal UpToDate As Boolean, ByVal NoRes As Boolean)
'*************************************************
'Author: BrianPr
'Created: 25/11/2008
'Last modified: 25/11/2008
'
'*************************************************
On Error GoTo error
    Dim extraArgs As String
    If Not UpToDate Then
        'No recibe update, ejecutar AU
        'Ejecuto el AoUpdate, sino me voy
        If Dir(App.path & "\AoUpdate.exe", vbArchive) = vbNullString Then
            MsgBox "No se encuentra el archivo de actualizaci�n AoUpdate.exe por favor descarguelo y vuelva a intentar", vbCritical
            End
        Else
            FileCopy App.path & "\AoUpdate.exe", App.path & "\AoUpdateTMP.exe"
            
            If NoRes Then
                extraArgs = " /nores"
            End If
            
            Call ShellExecute(0, "Open", App.path & "\AoUpdateTMP.exe", App.EXEName & ".exe" & extraArgs, App.path, SW_SHOWNORMAL)
            End
        End If
    Else
        If FileExist(App.path & "\AoUpdateTMP.exe", vbArchive) Then Kill App.path & "\AoUpdateTMP.exe"
    End If
Exit Sub

error:
    If Err.number = 75 Then 'Si el archivo AoUpdateTMP.exe est� en uso, entonces esperamos 5 ms y volvemos a intentarlo hasta que nos deje.
        Sleep 5
        Resume
    Else
        MsgBox Err.Description & vbCrLf, vbInformation, "[ " & Err.number & " ]" & " Error "
        End
    End If
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
    
    ClientSetup.bNoMusic = Not Audio.MusicActivated
    ClientSetup.bNoSound = Not Audio.SoundActivated
    ClientSetup.bNoSoundEffects = Not Audio.SoundEffectsActivated
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
    ListaClases(eClass.Minero) = "Minero"
    ListaClases(eClass.Herrero) = "Herrero"
    ListaClases(eClass.Experto_Madera) = "Experto en uso de madera"
    ListaClases(eClass.Talador) = "Le�ador"
    ListaClases(eClass.Carpintero) = "Carpintero"
    ListaClases(eClass.Pescador) = "Pescador"
    ListaClases(eClass.Sastre) = "Sastre"
    ListaClases(eClass.Alquimista) = "Alquimista"
    ListaClases(eClass.Luchador) = "Luchador"
    ListaClases(eClass.Con_Mana) = "Con uso de mana"
    ListaClases(eClass.Hechicero) = "Hechicero"
    ListaClases(eClass.Mago) = "Mago"
    ListaClases(eClass.Nigromante) = "Nigromante"
    ListaClases(eClass.Orden_Sagrada) = "Orden sagrada"
    ListaClases(eClass.Paladin) = "Paladin"
    ListaClases(eClass.Clerigo) = "Clerigo"
    ListaClases(eClass.Naturalista) = "Naturalista"
    ListaClases(eClass.Bardo) = "Bardo"
    ListaClases(eClass.Druida) = "Druida"
    ListaClases(eClass.Sigiloso) = "Sigiloso"
    ListaClases(eClass.Asesino) = "Asesino"
    ListaClases(eClass.Cazador) = "Cazador"
    ListaClases(eClass.Sin_Mana) = "Sin uso de mana"
    ListaClases(eClass.Arquero) = "Arquero"
    ListaClases(eClass.Guerrero) = "Guerrero"
    ListaClases(eClass.Caballero) = "Caballero"
    ListaClases(eClass.Bandido) = "Bandido"
    ListaClases(eClass.Pirata) = "Pirata"
    ListaClases(eClass.Ladron) = "Ladron"
    
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
    frmMain.RecTxt.Text = vbNullString
    
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
    
    Call Resolution.ResetResolution
    
    'Stop tile engine
    Call DeinitTileEngine
    
    Call SaveClientSetup
    
    'Destruimos los objetos p�blicos creados
    Set CustomMessages = Nothing
    Set CustomKeys = Nothing
    Set SurfaceDB = Nothing
    Set Dialogos = Nothing
  '  Set DialogosClanes = Nothing
    Set Audio = Nothing
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

Public Sub checkText(ByVal Text As String)
Dim Nivel As Integer
If Right(Text, Len(MENSAJE_FRAGSHOOTER_TE_HA_MATADO)) = MENSAJE_FRAGSHOOTER_TE_HA_MATADO Then
    Call ScreenCapture(True)
    Exit Sub
End If
If Left(Text, Len(MENSAJE_FRAGSHOOTER_HAS_MATADO)) = MENSAJE_FRAGSHOOTER_HAS_MATADO Then
    EsperandoLevel = True
    Exit Sub
End If
If EsperandoLevel Then
    If Right(Text, Len(MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA)) = MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA Then
        If CInt(mid(Text, Len(MENSAJE_FRAGSHOOTER_HAS_GANADO), (Len(Text) - (Len(MENSAJE_FRAGSHOOTER_PUNTOS_DE_EXPERIENCIA) + Len(MENSAJE_FRAGSHOOTER_HAS_GANADO))))) / 2 > ClientSetup.byMurderedLevel Then
            Call ScreenCapture(True)
        End If
    End If
End If
EsperandoLevel = False
End Sub

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

On Error GoTo Errhandler

    Dim nfile As Integer
    nfile = FreeFile ' obtenemos un canal
    Open App.path & "\errores.log" For Append Shared As #nfile
    Print #nfile, Date & " " & time & " " & Desc
    Close #nfile
    
    Exit Sub

Errhandler:

End Sub
