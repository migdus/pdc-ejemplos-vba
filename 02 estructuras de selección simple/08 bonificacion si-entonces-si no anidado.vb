' Autor: Miguel Dussán
 
' Enunciado del problema:
 
' Escriba un programa que determine la bonificación y el total que se
' le debe pagar a un empleado. Si el empleado gana máximo 1 millón de
' pesos, la bonificación será del 5% de su sueldo. Para sueldos superiores
' a esta cantidad, la compañía realiza unos descuentos para ayudar a los niños
' con problemas: si el empleado gana más de 1 millón de pesos y menos de 2 millones
' de pesos, se le hace una deducción del 0.5%; para montos superiores a este rango,
' la deducción será del 0.9%.
 
Option Explicit
 
Sub main()
 
    ' Variable del sueldo. Como es un entero largo, se utiliza
   ' Long en lugar de Integer
   
    Dim sueldo As Long
   
    'Variable para almacenar la bonificación. Como es un decimal
   ' grande, se utiliza Double en lugar de Float.
   
    Dim bono As Double
   
    ' Guarda el valor final entregado al empleado
   
    Dim valor_entregado As Double
   
    ' Acumula los resultados como cadena
   
    Dim cadena_resultado As String
   
    bono = 0 ' Si el sueldo es mayor a 1 millón, no hay bono
   
   
    ' La función Val(parámetro) convierte la cadena que tiene como
   ' parámetro en el valor numérico apropiado para la variable a la cual
   ' se le está haciendo la asignación.
   
    sueldo = Val(InputBox("Digite el sueldo del empleado"))
   
    cadena_resultado = "El sueldo ingresado fue $" & sueldo
   
    If sueldo <= 1000000 Then   ' Si sueldo <= 1000000 entonces
       bono = sueldo * 0.05
       
        cadena_resultado = cadena_resultado & ". El bono es de: $" & bono
        
        valor_entregado = sueldo + bono
   
    Else                        ' Si no
        ' Si no cumple la primera condición, significa que el empleado gana más de
        ' 1 millón de pesos. Para ese caso, tenemos un nuevo si-entonces, que está anidado en el lado
        ' del Si no de la primera condición.
        
        If sueldo < 2000000 Then ' Si entonces anidado
           bono = sueldo * 0.005
        Else
            bono = sueldo * 0.009
        End If ' Fin si entonces anidado
       
        cadena_resultado = cadena_resultado & ". Se le deduce: $" & bono
        
        ' Como para cualquiera de los dos casos el sueldo es mayor a un millón, entonces
        ' el bono no se le suma si no que se le resta al sueldo
        
        valor_entregado = sueldo - bono
       
    End If ' Fin si
   
    ' Se utiliza nuevamente cadena_resultado para tomar la cadena anterior y
   ' concatenarla con la nueva, que es la información del bono
   
    cadena_resultado = cadena_resultado & ". El total entregado de: $" & valor_entregado
   
    ' Mostrar la cadena que tiene todos los resultados acumulados
   MsgBox (cadena_resultado)
 
End Sub