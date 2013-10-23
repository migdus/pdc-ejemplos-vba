' Autor: Miguel Dussán

' Algoritmo de ordenamiento burbuja ascendente con arreglos

Option Explicit

Sub main()
    ' i : variable utilizada para iterar por el arreglo
    ' tam : cantidad de elementos a guardar en el arreglo
    ' arreglo : el arreglo
    
    Dim i, tam, temp, arreglo() As Integer
    
    'B andera que indica si se hizo al menos un intercambio durante el recorrido del arreglo.
    Dim hubo_intercambio As Boolean
    
    Do
        tam = Val(InputBox("¿Cantidad de números?"))
    Loop While tam < 1
    
    ReDim arreglo(tam) ' redimensionar el arreglo
    
    ' Captura de valores
    For i = 1 To tam
        arreglo(i) = Val(InputBox("Número para posición " & i & "?"))
    Next
    
    ' Hacer mientras se haya hecho intercambio
    Do
        ' El ciclo arranca asumiendo que no va a haber intercambios
        hubo_intercambio = False
        
        ' Se recorre todo el arreglo
        For i = 1 To tam
        
            ' Se verifican dos cosas:
            ' 1. Que exista una posición a continuación de la
            ' posición i-ésima, es decir, que exista una posición en
            ' la ubicación i + 1. Esto se verifica con la primera parte
            ' de la condición, que involucra el límite del arreglo, pues
            ' no existe un elemento i + 1 mayor o igual al tamaño del arreglo.

            ' 2. Que el contenido del arreglo en la posición actual sea mayor al
            ' contenido en la posición siguiente.
            
            If i + 1 <= tam Then
                If arreglo(i) > arreglo(i + 1) Then
                    ' Se realiza el intercambio
                    temp = arreglo(i + 1)
                    ' Se guarda el valor actual en la siguiente posición
                    arreglo(i + 1) = arreglo(i)
                    ' Se guarda el valor de la variable temporal en la posición actual
                    arreglo(i) = temp
                    ' Se marca la bandera en verdadero porque se hizo un intercambio
                    hubo_intercambio = True
                End If
            End If
        Next
        ' Si al terminar el recorrido por el arreglo no hubo intercambios,
        ' significa que el arreglo se encuentra organizado. En ese caso,
        ' la bandera debe encontrarse en falso, y el do-while termina.
    Loop While hubo_intercambio = True
    
    ' Se muestra el resultado
    
    Dim cadena_salida As String
    cadena_salida = "Arreglo ordenado: "
    For i = 1 To tam
        cadena_salida = cadena_salida & arreglo(i) & " "
    Next
    
    MsgBox (cadena_salida)
    
End Sub
