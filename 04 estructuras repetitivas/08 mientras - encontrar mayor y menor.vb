'
'  Autor: Miguel Dussán
'
'  Diseñe un programa que dados N números enteros ingresados por el usuario,
'  encuentre el mayor y el menor de estos datos.
'
'  Desarrollo:
'
'  Este ejercicio debe preguntar por la cantidad de números que debe capturar,
'  y para cada número capturado debe verificar si corresponde
'  al mayor o menor de los que el usuario ha digitado.
'
'  La estrategia de resolución involucrará a una bandera que tiene una función
'  especial: marcar si se ha ejecutado o no un bloque de código. Con esta
'  herramienta, dentro del ciclo se incluye un bloque si-entonces que se ejecutará
'  una sola vez, y dentro de él se inicializarán las variables que tienen los valores
'  menor y mayor; luego de esto se cambiará el valor de la bandera para que este bloque
'  no se ejecute nunca más.
'
'  EJERCICIOS
'
'  1) Modifique el programa para que el programa termine si el usuario indica que quiere
'  ingresar menos de cero números

Option Explicit

Sub main()

' Variables:
'
' a: Cantidad de números que el usuario desea ingresar.
' n: Número que el usuario ingresa para cada iteración del ciclo
' mayor: Almacena el número más grande de los ingresados por el usuario
' menor: Almacena el número más pequeño de los ingresados por el usuario

Dim n, a, c, mayor, menor As Integer

c = 0 ' Variable de control de ciclo

Dim b As Boolean
b = True ' Bandera para determinar si se ha fijado un valor inicial para las variables menor y mayor

a = Val(InputBox("¿Cuántos números desea ingresar?"))

Do While c < a
    n = Val(InputBox("¿Número?"))
    
    ' Este si-entonces verifica que sea la primera vez que se realiza una
    ' iteración. En caso de que así sea, se da a las variables mayor y menor
    ' como valor inicial el número que el usuario digitó. De esta manera, si ingresa
    ' más valores, ya se podrán almacenar correctamente como el menor o el mayor.
    '
    ' Al final se cambia el valor de la bandera a Falso para que no vuelva a realizar esa
    ' operación.
    
    If b = True Then
        mayor = n
        menor = n
        b = False
    Else
        ' Ejecuta esto a partir de la segunda iteración
        
        ' Si el número ingresado es más grande al número almacenado en la variable
        ' mayor, entonces este es el nuevo número mayor
        
        If n > mayor Then
            mayor = n
        Else
            ' Si el número ingresado es más pequeño que el almacenado en la variable menor, entonces
            ' este es el nuevo número menor
            
            If n < menor Then
                menor = n
            End If
        End If
    End If
    c = c + 1 ' Incremento de la variable de control de ciclo
Loop

' Mostrar los resultados
MsgBox ("Número menor: " & menor & ". Número mayor: " & mayor)
End Sub
