Option Explicit

' Autor: Miguel Dussán

' Programa que muestra los números del 10 al 0 de forma descendente

Sub main25()

    Dim numero As Integer
    Dim cadena_resultado As String
    
    cadena_resultado = ""
    numero = 10

    Do While numero > 0 ' Abre Hacer mientras
        
        cadena_resultado = cadena_resultado & numero
        
        numero = numero - 1
        
    Loop ' Cierra Hacer mientras
    
    MsgBox ("Resultado: " & acumulador)
End Sub