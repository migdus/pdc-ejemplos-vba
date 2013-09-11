' Autor: Miguel Dussán
 
' Enunciado del problema:
 
' Escriba un programa que determine la bonificación y el total que se
' le debe pagar a un empleado. Si el empleado gana a lo más 1 millón de
' pesos, la bonificación será del 5% de su sueldo. De lo contrario,
' del sueldo se le deducirá el 0.5% para ayudar a una fundación que trabaja
' con niños con problemas.
 
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
    
    ' Se concatena el resultado para este caso en la cadena que se va a
    ' mostrar al final. Tenga en cuenta que se concatena el contenido anterior
    ' con el texto y el valor del bono antes de asignar esta información a la variable
    ' cadena_resultado. Si no se hiciera esto, entonces el valor anterior de la variable se
    ' eliminaría y cambiaría únicamente por el valor del bono.
    
    cadena_resultado = cadena_resultado & ". El bono es de: $" & bono
    
    ' Como el porcentaje es un bono, se suma al sueldo y se almacena en la variable
    ' valor entregado
    
    valor_entregado = sueldo + bono
   
Else                        ' Si no
   bono = sueldo * 0.005   ' Deducible. Podemos usar la misma variable de bonificación
   
   ' Se concatean el resultado pero esta vez no se dice que es un bono, si no que es un deducible
   
   cadena_resultado = cadena_resultado & ". Se le deduce: $" & bono
   
   
     ' Como el porcentaje es un deducible, se resta al sueldo y se almacena en la variable
    ' valor entregado
    
    valor_entregado = sueldo - bono
   
End If ' Fin si
 
' Se utiliza nuevamente cadena_resultado para tomar la cadena que ya se venía
' acumulando, y se concatena con el valor entregado al empleado
 
cadena_resultado = cadena_resultado & ". El total entregado de: $" & valor_entregado
 
' Mostrar la cadena que tiene todos los resultados acumulados

MsgBox (cadena_resultado)
 
End Sub
