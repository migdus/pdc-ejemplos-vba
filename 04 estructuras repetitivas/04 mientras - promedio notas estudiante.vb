Option Explicit

' Autor: Miguel Dussán

' Programa que captura notas de un alumno hasta que el usuario introduce un número negativo.
' En ese instante calcula y muestra el promedio

Sub main()

    Dim nota, acumulador, contador As Integer
    
    nota = 0
    
    ' Variable acumuladora
    ' Almacena resultados parciales de la suma
    
    acumulador = 0
    
    ' Variable que lleva la cuenta de la cantidad de notas ingresadas
    
    contador = 0
    
    Do While nota >= 0 ' Abre Hacer mientras
        
        ' Pide una nota
        nota = Val(InputBox("Nota? "))
        
        ' Si el número ingresado no es negativo, entonces es una nota válida
        If nota >= 0 Then
            ' Se acumula la nota
            acumulador = acumulador + nota
            contador = contador + 1 ' Se incrementa el contador de notas ingresadas
        End If
        
    Loop ' Cierra Hacer mientras
    
    ' Una vez terminado el ciclo se calcula el promedio y se muestra
    
    MsgBox ("Promedio de notas: " & acumulador / contador)
End Sub
