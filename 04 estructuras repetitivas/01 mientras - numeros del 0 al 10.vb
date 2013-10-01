Option Explicit

' Autor: Miguel Dussán

' Programa que muestra los números desde 0 hasta 10 utilizando ciclos

Sub main()

    ' Variable de control de ciclo
    ' Se usa también para mostrar el número actual
    
    Dim numero As Integer
    
    ' Acumula los resultados para mostrarlos al final en un solo
    ' MsgBox
    
    Dim cadena_resultado As String
    
    cadena_resultado = "" ' Inicialización con cadena vacía
    numero = 0 ' Inicialización
    
    Do While numero <= 10 ' Abre Hacer mientras
    
    ' Se concatena el número a la cadena de resultado
    cadena_resultado = cadena_resultado & numero & " "
    
    numero = numero + 1 ' Incremento de la variable de control
    Loop ' Cierra Hacer mientras
    
    MsgBox (cadena_resultado)

End Sub