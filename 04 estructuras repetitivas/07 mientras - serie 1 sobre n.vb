' Autor: Miguel Dussán
' Diseñe un programa que dado un número entero positivo N calcule el resultado de la serie: 1 - 1/2 + 1/4 - ... +- 1/n
'
' Desarrollo:
'
' El enunciado muestra que:
' 1) El número que el usuario ingresa determina el término máximo de la serie.
' 2) Se alternan los símbolos de operación: suma y resta, arrancando por la substracción
' Se utilizará una bandera para alternar la operación que se realiza. Los resultados se almacenarán en una variable
' acumuladora.

Option Explicit

Sub main()

Dim n As Long ' Número hasta el cual irá la serie

Dim c, acum As Double ' Variable de control de ciclo inicializada en el primer término de la serie
c = 1
acum = 0

Dim b As Boolean
b = True ' Bandera inicializada en Verdadero.

n = Val(InputBox("¿Número límite de la serie?"))

Do While c <= n ' Mientras no se llegue al número del término dado por el usuario
    If b = True Then
        
        ' Si la bandera es verdadera, se suma el término.
        ' El valor del término de la serie está dado por 1/c, donde c
        ' es la variable de control.
        '
        ' En la primera iteración se acumula el valor del primer término, que es 1/1
        
        acum = acum + 1 / c
        b = False
    Else
        acum = acum - 1 / c
        b = True
    End If
    
    c = c + 1 ' incremento de la variable de control
    
Loop

MsgBox ("El resultado de la operación de la serie es " & acum)

End Sub
