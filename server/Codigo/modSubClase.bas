Attribute VB_Name = "modSubClase"
Option Explicit

Public Sub EnviarSubClase(ByVal UserIndex As Integer)

    If UserList(UserIndex).flags.Muerto = 0 Then
        Call WriteShowClaseForm(UserIndex, UserList(UserIndex).Clase)
    End If
            
End Sub

Public Sub RecibirSubClase(ByVal UserIndex As Integer, ByVal Clase As Byte)

    If Not PuedeSubirClase(UserIndex) Then Exit Sub

    With UserList(UserIndex)

        Select Case .Clase

            Case eClass.Ciudadano
                If Clase = 1 Then
                    .Clase = eClass.Trabajador
                Else
                    .Clase = eClass.Luchador
                End If
            
            Case eClass.Trabajador
                Select Case Clase
                
                    Case 1: .Clase = eClass.Experto_Minerales
                    Case 2: .Clase = eClass.Experto_Madera
                    Case 3: .Clase = eClass.Pescador
                    Case 4: .Clase = eClass.Sastre
                
                End Select
            
            Case eClass.Experto_Minerales
                If Clase = 1 Then
                    .Clase = eClass.Minero
                Else
                    .Clase = eClass.Herrero
                End If
            
            Case eClass.Experto_Madera
                If Clase = 1 Then
                    .Clase = eClass.Talador
                Else
                    .Clase = eClass.Carpintero
                End If
            
            Case eClass.Luchador
                If Clase = 1 Then
                    .Clase = eClass.Con_Mana
                    Call AprenderHechizo(UserIndex, 2)
                    .Stats.MaxMAN = 100
                    Call WriteUpdateMana(UserIndex)
                    If Not PuedeSubirClase(UserIndex) Then Call WriteSubeClase(UserIndex, False)
                Else
                    .Clase = eClass.Sin_Mana
                End If
            
            Case eClass.Con_Mana
                Select Case Clase
                
                    Case 1: .Clase = eClass.Hechicero
                    Case 2: .Clase = eClass.Orden_Sagrada
                    Case 3: .Clase = eClass.Naturalista
                    Case 4: .Clase = eClass.Sigiloso
                    
                
                End Select
            
            Case eClass.Hechicero
                If Clase = 1 Then
                    .Clase = eClass.Mago
                Else
                    .Clase = eClass.Nigromante
                End If
            
            Case eClass.Orden_Sagrada
                If Clase = 1 Then
                    .Clase = eClass.Paladin
                Else
                    .Clase = eClass.Clerigo
                End If
            Case eClass.Naturalista
                If Clase = 1 Then
                    .Clase = eClass.Bardo
                Else
                    .Clase = eClass.Druida
                End If
            
            Case eClass.Sigiloso
                If Clase = 1 Then
                    .Clase = eClass.Asesino
                Else
                    .Clase = eClass.Cazador
                End If
            
            Case eClass.Sin_Mana
                If Clase = 1 Then
                    .Clase = eClass.Bandido
                Else
                    .Clase = eClass.Caballero
                End If
            
            Case eClass.Bandido
                If Clase = 1 Then
                    .Clase = eClass.Pirata
                Else
                    .Clase = eClass.Ladron
                End If
            
            Case eClass.Caballero
                If Clase = 1 Then
                    .Clase = eClass.Guerrero
                Else
                    .Clase = eClass.Arquero
                End If
        End Select

    End With

Call CalcularValores(UserIndex)
If Not PuedeSubirClase(UserIndex) Then Call WriteSubeClase(UserIndex, False)

End Sub
