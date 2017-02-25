Attribute VB_Name = "Carteles"
'Argentum Online 0.11.6
'
'Copyright (C) 2002 Márquez Pablo Ignacio
'Copyright (C) 2002 Otto Perez
'Copyright (C) 2002 Aaron Perkins
'Copyright (C) 2002 Matías Fernando Pequeño
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
'Calle 3 número 983 piso 7 dto A
'La Plata - Pcia, Buenos Aires - Republica Argentina
'Código Postal 1900
'Pablo Ignacio Márquez

Option Explicit

Const XPosCartel = 360
Const YPosCartel = 335
Const MAXLONG = 40

'Carteles
Public Cartel As Boolean
Public Leyenda As String
Public LeyendaFormateada() As String
Public textura As Integer


Sub InitCartel(Ley As String, Grh As Integer)
If Not Cartel Then
    Leyenda = Ley
    textura = Grh
    Cartel = True
    ReDim LeyendaFormateada(0 To (Len(Ley) \ (MAXLONG \ 2)))
                
    Dim I As Integer, k As Integer, anti As Integer
    anti = 1
    k = 0
    I = 0
    Call DarFormato(Leyenda, I, k, anti)
    I = 0
    Do While LeyendaFormateada(I) <> "" And I < UBound(LeyendaFormateada)
        
       I = I + 1
    Loop
    ReDim Preserve LeyendaFormateada(0 To I)
Else
    Exit Sub
End If
End Sub


Private Function DarFormato(s As String, I As Integer, k As Integer, anti As Integer)
If anti + I <= Len(s) + 1 Then
    If ((I >= MAXLONG) And mid$(s, anti + I, 1) = " ") Or (anti + I = Len(s)) Then
        LeyendaFormateada(k) = mid(s, anti, I + 1)
        k = k + 1
        anti = anti + I + 1
        I = 0
    Else
        I = I + 1
    End If
    Call DarFormato(s, I, k, anti)
End If
End Function


Sub DibujarCartel()
If Not Cartel Then Exit Sub
Dim X As Integer, Y As Integer
X = XPosCartel + 20
Y = YPosCartel + 60
'Call DDrawTransGrhIndextoSurface(textura, XPosCartel, YPosCartel, 0, AmbientColor)
Dim J As Integer, desp As Integer

For J = 0 To UBound(LeyendaFormateada)
    'Text_Draw X, Y + desp, LeyendaFormateada(J), vbWhite
    desp = desp + (frmMain.Font.size) + 5
Next
End Sub

