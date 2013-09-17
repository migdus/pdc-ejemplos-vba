' Autor: Miguel Dussán

' Programa que calcula la suma de los números entre 0 y 10

Option Explicit

Sub main()

    Dim numero, acumulador As Integer
    
    numero = 0
    
    ' Variable acumuladora
    ' Almacena resultados parciales de la suma
    
    acumulador = 0
    
    Do While numero <= 10 ' Abre Hacer mientras
        
        acumulador = acumulador + numero
        
        numero = numero + 1
        
    Loop ' Cierra Hacer mientras
    
    MsgBox ("Resultado: " & acumulador)
End Sub