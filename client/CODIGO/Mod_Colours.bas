Attribute VB_Name = "Mod_Colours"
'Some awesome colours would be placed here in some time
Option Explicit

Public White(3) As Long
Public Red(3) As Long
Public Cyan(3) As Long
Public Black(3) As Long
Public Yellow(3) As Long

Public Sub InitColours()
        
    White(0) = D3DColorXRGB(255, 255, 255)
    White(1) = White(0)
    White(2) = White(0)
    White(3) = White(0)
    
    Red(0) = D3DColorXRGB(255, 0, 0)
    Red(1) = Red(0)
    Red(2) = Red(0)
    Red(3) = Red(0)
    
    Cyan(0) = D3DColorXRGB(0, 255, 255)
    Cyan(1) = Cyan(0)
    Cyan(2) = Cyan(0)
    Cyan(3) = Cyan(0)
    
    Black(0) = 0
    Black(1) = 0
    Black(2) = 0
    Black(3) = 0
    
    Yellow(0) = D3DColorXRGB(255, 255, 0)
    Yellow(1) = Yellow(0)
    Yellow(2) = Yellow(0)
    Yellow(3) = Yellow(0)
    
End Sub
