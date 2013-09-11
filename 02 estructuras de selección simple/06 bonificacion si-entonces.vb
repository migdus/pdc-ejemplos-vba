' Autor: Miguel Dussán

' Enunciado del problema:

' Escriba un programa que determine la bonificación y el total que se
' le debe pagar a un empleado. Si el empleado gana a lo más 1 millón de
' pesos, la bonificación será del 5% de su sueldo. De lo contrario no
' tendrá bonificación.

Option Explicit

Sub main()

    ' Variable del sueldo. Como es un entero largo, se utiliza
    ' Long en lugar de Integer
    Dim sueldo As Long
    
    ' Variable para almacenar la bonificación. Como es un decimal
    ' grande, se utiliza Double en lugar de Float.
    Dim bono As Double
    
    ' Guarda el valor final entregado al empleado
    Dim valor_entregado As Double
    
    
    bono = 0 ' Si el sueldo es mayor a 1 millón, no hay bono
    
    ' La función Val(parámetro) convierte la cadena que tiene como
    ' parámetro en el valor numérico apropiado para la variable a la cual
    ' se le está haciendo la asignación.
    
    sueldo = Val(InputBox("Digite el sueldo del empleado"))
    
    If sueldo <= 1000000 Then ' Si sueldo <= 1000000 entonces
        bono = sueldo * 0.05
    End If ' Fin si
    
    valor_entregado = sueldo + bono
    
    ' Guarda los resultados como cadena y después los muestra
    
    Dim cadena_resultado As String
    
    cadena_resultado = "El sueldo ingresado fue $" & sueldo
    
    ' Se utiliza nuevamente cadena_resultado para tomar la cadena anterior y
    ' concatenarla con la nueva, que es la información del bono
    
    cadena_resultado = cadena_resultado & ". El bono es de: $" & bono
    cadena_resultado = cadena_resultado & ". El total entregado de: $" & valor_entregado
    
    ' Mostrar la cadena que tiene todos los resultados
    MsgBox (cadena_resultado)

End Sub