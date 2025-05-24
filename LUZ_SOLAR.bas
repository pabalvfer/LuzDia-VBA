Attribute VB_Name = "LUZ_SOLAR"
'AUTOR: PABLO ALVAREZ FERNANDEZ
'VERSION: 1.1
'FECHA REVISIÓN: 04/02/2025



'FUNCIONES DE MODULO
'    Function DeclinacionSolar(fecha As Variant, Optional ecuacion As Integer = SPENCER) As Double
'    Function HoraOrto(dia As Variant, latitud As Double, longitud As Double, Optional outPutFormat As Integer = HORA_NUMERIC) As Double
'    Function HoraOcaso(dia As Variant, latitud As Double, longitud As Double, Optional outPutFormat As Integer = HORA_NUMERIC) As Double
'    Function HorasLuzDia(dia As Variant, latitud As Double) As Double
'    Function diaJuliano(fecha As Variant) As Integer
'    Function AnguloDiario(diaJuliano As Double, bisiesto As Boolean, Optional hora As Double = 12) As Double
'    Function EcuacionTiempo(fecha As Variant) As Double
'    Function HorasDeg2Reloj(HoraDeg As Double, fecha As Variant, longitud As Double, Optional timezone = UTC_1) As Double
'    Function IsLeapYear(year As Integer) As Boolean



'ESTE MODULO UTILIZA ESTAS FUNCIONES DEL MÓDULO DE TRIGONOMETRIA
'    Function pi() As Double
'    Function Deg2Rad(Deg As Double) As Double
'    Function Rad2Deg(Rad As Double) As Double
'    Function ArcCos(Rad As Double) As Double
'    Function ArcSin(Rad As Double) As Double

Option Explicit


' Constantes para las ecuaciones de declinación solar
Public Const SPENCER As Integer = 1
Public Const COOPER As Integer = 2

' Constante para el formato de salida de horas
Public Const HORA_NUMERIC As Integer = 1 ' Hora en formato numerico (ej. 14.5 para 14:30)
Public Const HORA_GRADOS As Integer = 2 ' Hora en grados solares

' Constante para la zona horaria (UTC+1 para España)
Public Const UTC_1 As Integer = 1 ' Esta constante solo representa el offset de 1 hora, no el cambio de horario de verano/invierno



Function DeclinacionSolar(fecha As Variant, Optional ecuacion As Integer = SPENCER) As Double
    ' Devuelve la declinación solar según fórmula de Cooper (1965) o Spencer (1971) en radianes.
    ' SPENCER: ecuación = 1 (más precisa)
    ' COOPER:  ecuación = 2
    ' Si es un número, lo considera como día juliano (NO ADMITE BISIESTOS para la entrada numérica).
    ' dia = 1 para el 1 de enero
    ' dia = 365 para el 31 de diciembre
    
    Dim dia As Double
    Dim angDia As Double
    Dim bisiesto As Boolean

    DeclinacionSolar = -2
    
    'VALIDACIONES
    'Detecta si fecha es una fecha (Date) o un dia juliano (numerico)
    If IsDate(fecha) Then
        dia = diaJuliano(fecha)
        bisiesto = IsLeapYear(year(fecha))
    ElseIf IsNumeric(fecha) Then
        ' Para entrada numérica, se asume un año no bisiesto
        dia = fecha
        bisiesto = False
    Else
        Exit Function
    End If
        
    If dia <= 0 Or dia > 365 + CInt(bisiesto) Then Exit Function
          
    'Calcula la declinacion solar
    If ecuacion = SPENCER Then
        angDia = AnguloDiario(dia, bisiesto)
        DeclinacionSolar = 0.006918 - 0.399912 * Cos(angDia) + _
                           0.070257 * Sin(angDia) - 0.006758 * Cos(2 * angDia) + _
                           0.000907 * Sin(2 * angDia) - 0.002697 * Cos(3 * angDia) + _
                           0.00148 * Sin(3 * angDia)
    ElseIf ecuacion = COOPER Then
        DeclinacionSolar = 23.45 * Sin(Deg2Rad(360 / (365 + CInt(bisiesto)) * (dia + 284)))
    End If
                   
End Function

Private Function AnguloHorario(dia As Variant, latitud As Double, longitud As Double, Optional ecuacion As Integer = SPENCER) As Double
    ' Función auxiliar para el calculo del angulo horario para del amanecer y aterdecer
    ' dia: orden del día del año o fecha.
    ' latitud: Latitud en grados.
    ' longitud: Longitud en grados.

    ' ecuacion:define la formala para el calculo de la declinación solar
    '   SPENCER: ecuación = 1 (más precisa)
    '   COOPER:  ecuación = 2


    Dim DecSolar As Double
    Dim HoraGrados As Double

    'Calcula la declinación
    DecSolar = DeclinacionSolar(dia, ecuacion)
        
    If DecSolar = -2 Then Exit Function
    
    'Calcula la hora en grados

    AnguloHorario = ArcCos(Cos(Deg2Rad(90.833)) / (Cos(Deg2Rad(latitud)) / Cos(DecSolar)) - Tan(Deg2Rad(latitud)) * Tan(DecSolar))
       
    Exit Function

End Function

Function HoraOrto(dia As Variant, latitud As Double, longitud As Double, Optional outPutFormat As Integer = HORA_NUMERIC, Optional ecuacion As Integer = SPENCER) As Double
    ' Devuelve la hora del orto (amanecer) expresada en grados u hora decimal.
    ' LA HORA EN GRADOS ES HORA SOLAR Y LA DECIMAL ES HORA LOCAL.
    ' dia: orden del día del año o fecha.
    ' latitud: Latitud en grados.
    ' longitud: Longitud en grados.
    ' outPutFormat: define el formato de salida
    ' ecuacion:define la formala para el calculo de la declinación solar
    '   SPENCER: ecuación = 1 (más precisa)
    '   COOPER:  ecuación = 2

    Dim HoraGrados As Double
    Dim HoraRad As Double

    'Calcula la hora en grados
    HoraRad = AnguloHorario(dia, latitud, longitud, ecuacion)
    HoraGrados = Rad2Deg(HoraRad)
    
    'Devuelve la hora en el formato solicitado
    If outPutFormat = HORA_NUMERIC Then
        HoraOrto = HorasDeg2Reloj(HoraGrados, dia, longitud)
    Else
        HoraOrto = HoraGrados
    End If

End Function


Function HoraOcaso(dia As Variant, latitud As Double, longitud As Double, Optional outPutFormat As Integer = HORA_NUMERIC, Optional ecuacion As Integer = SPENCER) As Double
    ' Devuelve la hora del orto (amanecer) expresada en grados u hora decimal.
    ' LA HORA EN GRADOS ES HORA SOLAR Y LA DECIMAL ES HORA LOCAL.
    ' dia: orden del día del año o fecha.
    ' latitud: Latitud en grados.
    ' longitud: Longitud en grados.
    ' outPutFormat: define el formato de salida
    ' ecuacion:define la formala para el calculo de la declinación solar
    '   SPENCER: ecuación = 1 (más precisa)
    '   COOPER:  ecuación = 2

    Dim HoraGrados As Double
    Dim HoraRad As Double
    
    'Calcula la hora en grados
    HoraRad = AnguloHorario(dia, latitud, longitud, ecuacion)
    HoraGrados = -Rad2Deg(HoraRad)
    
    'Devuelve la hora en el formato solicitado
    If outPutFormat = HORA_NUMERIC Then
        HoraOcaso = HorasDeg2Reloj(HoraGrados, dia, longitud)
    Else
        HoraOcaso = HoraGrados
    End If

End Function



Function HorasLuzDia(dia As Variant, latitud As Double, Optional ecuacion As Integer = SPENCER) As Double
    ' Calcula las horas de luz solar del dia
    ' dia: orden del día del año o fecha.
    ' latitud: Latitud en grados.
    ' ecuacion:define la formala para el calculo de la declinación solar
    '   SPENCER: ecuación = 1 (más precisa)
    '   COOPER:  ecuación = 2

    Dim DecSolar As Double
    
    HorasLuzDia = -1
    
    'Calcula la declinación
    DecSolar = DeclinacionSolar(dia, ecuacion)
    If DecSolar = -2 Then Exit Function

    'Calcula el numero de horas de luz
    HorasLuzDia = Rad2Deg(ArcCos(Cos(Deg2Rad(90.833)) / (Cos(Deg2Rad(latitud)) / Cos(DecSolar)) - Tan(Deg2Rad(latitud)) * Tan(DecSolar))) * 2 / 15

End Function



Function diaJuliano(fecha As Variant) As Integer
    'Calcula el dia juliano de la fecha para ese año
    'OrdenDiaAno= 1 para el 1 de enero y 365 para el 31 de diciembre (366 si es bisiesto)

    Dim Primerdia
    
    diaJuliano = -1
    
    If IsDate(fecha) Then
       
            Primerdia = DateSerial(year(fecha), 1, 1)
            
            diaJuliano = DateDiff("d", Primerdia, fecha) + 1
            
    End If


End Function

Function AnguloDiario(diaJuliano As Double, bisiesto As Boolean, Optional hora As Double = 12) As Double
'Calcula el angulo diario en radianes a partir del dia juliano y la hora opcional

    Dim diasEnAno  As Integer
    
    If bisiesto Then
        diasEnAno = 366
    Else
        diasEnAno = 365
    End If
    
    AnguloDiario = 2 * pi / diasEnAno * (diaJuliano - 1 + (hora - 12) / 24)
        

End Function

Function EcuacionTiempo(fecha As Variant) As Double
    ' Devuelve la corrección del tiempo en minutos según Spencer (1971).

    ' Si fecha es un número, lo considera el número de día juliano (NO ADMITE BISIESTOS).
    ' dia = 1 para el 1 de enero
    ' dia = 365 para el 31 de diciembre


    Dim dia As Double
    Dim angDia As Double
    Dim bisiesto As Boolean
    
    EcuacionTiempo = -2
    
    'VALIDACIONES
    'Detecta si fecha es una fecha (Date) o un dia juliano (numerico)
    If IsDate(fecha) Then
        dia = diaJuliano(fecha)
        bisiesto = IsLeapYear(year(fecha))
    ElseIf IsNumeric(fecha) Then
        ' Para entrada numérica, se asume un año no bisiesto
        dia = fecha
        bisiesto = False
    Else
        Exit Function
    End If
        
    If dia <= 0 Or dia > 365 + CInt(bisiesto) Then Exit Function

    'calcula el angulo diario en radianes
    angDia = AnguloDiario(dia, bisiesto)
    
    'calcula la correccion del tiempo
    EcuacionTiempo = 229.18 * (0.000075 + 0.001868 * Cos(angDia) _
                     - 0.032077 * Sin(angDia) - 0.014615 * Cos(2 * angDia) _
                     - 0.040849 * Sin(2 * angDia))
        
    
End Function

Function HorasDeg2Reloj(HoraDeg As Double, fecha As Variant, longitud As Double, Optional timezone = UTC_1) As Double
'Convierte la hora solar expresada en grados a hora local en decimal
'NO TIENE EN CUENTA CAMBIO DE HORA VERANO/INVIERNO
'HoraDeg: hora solar en grados
'fecha: en formato fecha o dia juliano
'longitud: en grados
'timezone: zona UTC, por defecto España UTC+1

    Dim timeoffset As Double
    
    'Cacula el desfase por longitud y ecucion del tiempo
    timeoffset = EcuacionTiempo(fecha) + 4 * longitud  'resultado en minutos
    
    'Convierte a hora local
    HorasDeg2Reloj = (12 - HoraDeg / 15) - timeoffset / 60 + timezone

    
        
End Function

Function IsLeapYear(year As Integer) As Boolean
'Verifica si es un año bisiesto

    If IsLeapYear = (year Mod 4 = 0 And year Mod 100 <> 0) Or (year Mod 400 = 0) Then
       IsLeapYear = True
    Else
       IsLeapYear = False
    End If
 
End Function
