Attribute VB_Name = "modClases"
Option Explicit

Public Resta(1 To NUMCLASES) As Single
Public Recompensas() As Recompensa
Public AumentoHit(1 To NUMCLASES) As Byte

Sub EstablecerRestas()

Resta(eClass.Ciudadano) = 3
AumentoHit(eClass.Ciudadano) = 3

Resta(eClass.Trabajador) = 2.5
AumentoHit(eClass.Trabajador) = 3

Resta(eClass.Experto_Minerales) = 2.5
AumentoHit(eClass.Experto_Minerales) = 3

Resta(eClass.Minero) = 2.5
AumentoHit(eClass.Minero) = 2

Resta(eClass.Herrero) = 2.5
AumentoHit(eClass.Herrero) = 2

Resta(eClass.Experto_Madera) = 2.5
AumentoHit(eClass.Experto_Madera) = 3

Resta(eClass.Talador) = 2.5
AumentoHit(eClass.Talador) = 2

Resta(eClass.Carpintero) = 2.5
AumentoHit(eClass.Carpintero) = 2

Resta(eClass.Pescador) = 2.5
AumentoHit(eClass.Pescador) = 1

Resta(eClass.Sastre) = 2.5
AumentoHit(eClass.Sastre) = 2

Resta(eClass.Alquimista) = 2.5
AumentoHit(eClass.Alquimista) = 2

Resta(eClass.Luchador) = 3
AumentoHit(eClass.Luchador) = 3

Resta(eClass.Con_Mana) = 3
AumentoHit(eClass.Con_Mana) = 3

Resta(eClass.Hechicero) = 3
AumentoHit(eClass.Hechicero) = 3

Resta(eClass.Mago) = 3
AumentoHit(eClass.Mago) = 1

Resta(eClass.Nigromante) = 3
AumentoHit(eClass.Nigromante) = 1

Resta(eClass.Orden_Sagrada) = 1.5
AumentoHit(eClass.Orden_Sagrada) = 3

Resta(eClass.Paladin) = 0.5
AumentoHit(eClass.Paladin) = 3

Resta(eClass.Clerigo) = 1.5
AumentoHit(eClass.Clerigo) = 2

Resta(eClass.Naturalista) = 2.5
AumentoHit(eClass.Naturalista) = 3

Resta(eClass.Bardo) = 1.5
AumentoHit(eClass.Bardo) = 2

Resta(eClass.Druida) = 3
AumentoHit(eClass.Druida) = 2

Resta(eClass.Sigiloso) = 1.5
AumentoHit(eClass.Sigiloso) = 3

Resta(eClass.Asesino) = 1.5
AumentoHit(eClass.Asesino) = 3

Resta(eClass.Cazador) = 0.5
AumentoHit(eClass.Cazador) = 3

Resta(eClass.Sin_Mana) = 2
AumentoHit(eClass.Sin_Mana) = 2

AumentoHit(eClass.Arquero) = 3

AumentoHit(eClass.Guerrero) = 3

AumentoHit(eClass.Caballero) = 3

AumentoHit(eClass.Bandido) = 2

Resta(eClass.Pirata) = 1.5
AumentoHit(eClass.Pirata) = 2

Resta(eClass.Ladron) = 2.5
AumentoHit(eClass.Ladron) = 2

End Sub

Public Function ClaseBase(ByVal Clase As eClass) As Boolean

ClaseBase = (Clase = eClass.Ciudadano Or Clase = eClass.Trabajador Or Clase = eClass.Experto_Minerales Or _
            Clase = eClass.Experto_Madera Or Clase = eClass.Luchador Or Clase = eClass.Con_Mana Or _
            Clase = Hechicero Or Clase = Orden_Sagrada Or Clase = Naturalista Or _
            Clase = eClass.Sigiloso Or Clase = eClass.Sin_Mana Or Clase = eClass.Bandido Or _
            Clase = eClass.Caballero)

End Function

Public Function PuedeFaccion(ByVal UserIndex As Integer)
    
    With UserList(UserIndex)
    
        PuedeFaccion = Not EsNewbie(UserIndex) And _
        (.Faccion.BandoOriginal = eFaccion.Neutral) And _
        (.flags.Privilegios And PlayerType.User) 'and (.guilindex > 0)
    
    End With
End Function
Public Function PuedeSubirClase(ByVal UserIndex As Integer) As Boolean

    With UserList(UserIndex)
        PuedeSubirClase = (.Stats.ELV >= 3 And .Clase = eClass.Ciudadano) Or _
                    (.Stats.ELV >= 6 And (.Clase = eClass.Luchador Or .Clase = eClass.Trabajador)) Or _
                    (.Stats.ELV >= 9 And (.Clase = eClass.Experto_Minerales Or .Clase = eClass.Experto_Madera Or .Clase = eClass.Con_Mana Or .Clase = eClass.Sin_Mana)) Or _
                    (.Stats.ELV >= 12 And (.Clase = eClass.Caballero Or .Clase = eClass.Bandido Or .Clase = eClass.Hechicero Or .Clase = eClass.Naturalista Or .Clase = eClass.Orden_Sagrada Or .Clase = eClass.Sigiloso))
    
    End With
    
End Function
