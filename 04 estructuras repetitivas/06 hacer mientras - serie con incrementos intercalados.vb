utor: Miguel Dussán

' Escriba un programa que muestre los términos de la siguiente serie y calcule su suma:
' 2,5,7,10,12,15,17,... 1800

Option Explicit

Sub main()

'Variable de control de ciclo y acumulador
Dim control, acumulador As Long

'Variable bandera
'Permite seleccionar si se va a incrementar en 2 o en 3
'la variable de control de ciclo
Dim bandera As Boolean

'Cadena para acumular resultados y mostrarlos al final
Dim cadena_resultado As String

'Inicialización de cadena vacía
cadena_resultado = ""

'Inicialización de variable bandera
' Cuando esta sea Verdadera, se le sumará 3 a la variable de control
' De lo contrario se le sumará 2
bandera = True

control = 2 ' Inicialización en el primer número de la serie

acumulador = 0 ' Inicialización del acumulador

' Mientras la variable de control sea menor o igual
' al límite superior de la serie
Do While control <= 1800

' Acumula resultados
cadena_resultado = cadena_resultado & control & " "

' Acumula los valores de la serie
acumulador = acumulador + control

' Si la bandera es verdadera, incrementar en 3
If bandera = True Then

control = control + 3
bandera = False ' Cambiar valor de bandera

Else
' De otra forma, la bandera es falsa
control = control + 2 ' Incrementar en 2
bandera = True ' Cambiar valor de la bandera

End If

Loop

' El resultado se corta por la limitación de MsgBox

MsgBox (cadena_resultado)

' Resultado de la suma
MsgBox ("Suma: " & acumulador)
End Sub


