Attribute VB_Name = "TRIGONOMETRIA"
'Function pi() As Double
'Function Deg2Rad(Deg As Double) As Double
'Function Rad2Deg(Rad As Double) As Double
'Function ArcCos(Rad As Double) As Double
'Function ArcSin(Rad As Double) As Double
'Function ArcTan(Rad As Double) As Double




Function pi() As Double
'Devuelve el numero Pi

    pi = Application.WorksheetFunction.pi

End Function

Function Deg2Rad(Deg As Double) As Double
'Convierte grados a radianes

    Deg2Rad = Deg * pi / 180

End Function

Function Rad2Deg(Rad As Double) As Double
'Convierte radianes a grados

    Rad2Deg = Rad * 180 / pi

End Function


Function ArcCos(Rad As Double) As Double
'Devuelve el arcocoseno de un angulo expresado en radianes
    
        ArcCos = Application.WorksheetFunction.Acos(Rad)
    
End Function

Function ArcSin(Rad As Double) As Double
'Devuelve el arcoseno de un angulo expresado en radianes
    
        ArcSin = Application.WorksheetFunction.Asin(Rad)
    
End Function

Function ArcTan(Rad As Double) As Double
'Devuelve el arcotangente de un angulo expresado en radianes
    
        ArcTan = Application.WorksheetFunction.Atan(Rad)
    
End Function
