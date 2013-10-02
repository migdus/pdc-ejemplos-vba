' Autor: Miguel Dussán

' Escriba un programa que capture un número n, y a partir de este imprima una
' sucesión de caracteres tal que en la primera línea aparezca #, en la segunda línea ##,
' en la tercera ###, y así sucesivamente hasta llegar a la línea n, donde se imprimirán n caracteres #.

Option Explicit

Sub main()

' base, exponente: Argumentos de entrada
' cont_multip: contadora del ciclo anidado
' multip: resultado de una multiplicación
' cont_pot: contadora del cliclo exterior, que controla el número de multiplicaciones que se realizan
' pot: almacena el resultado de la potenciación
Dim base, exponente, cont_multip, multip, cont_pot, pot As Long

cont_pot = 0
pot = 1

base = Val(InputBox("Base?"))
exponente = Val(InputBox("Exponente?"))

Do While cont_pot < exponente
    
    multip = 0
    cont_multip = 0
    
    ' Hacer la multiplicación de la base con el resultado
    ' de la anterior multiplicación, que se almacena en pot
    
    Do While cont_multip < pot
        multip = multip + base ' se acumulan sumas sucesivas
        cont_multip = cont_multip + 1
    Loop
    
    pot = multip
    cont_pot = cont_pot + 1
Loop

MsgBox ("Resultado: " & pot)
End Sub
