 ' Autor: Miguel Dussán
 '
 '
 ' Diseñe un programa que calcule el término número n de la serie de Fibonacci de
 ' acuerdo a la elección del usuario. La serie de Fibonacci tiene como primero y
 ' segundo término 0 y 1, respectivamente, y cada término adicional de la serie
 ' corresponde a la suma de los dos números anteriores.
 '
 ' Ejemplo de la serie: 0, 1, 1, 2, 3, 5, 8, 13, 21, 34, 55, 89
 
 ' Ejercicios
 
 ' 1) Modifique el programa para que no solamente muestre el número de la serie que el
 '     usuario indique, si no también todos los números de la serie hasta ese número

Option Explicit

Sub main()

Dim num_termino As Integer ' Número del término que se debe encontrar
Dim control As Integer ' Variable de control del ciclo
Dim term_a, term_b As Integer ' Primer y segundo término de la serie
Dim temp As Integer

control = 3
term_a = 0
term_b = 1

num_termino = Val(InputBox("Ingrese el término que desea averiguar"))

If num_termino >= 3 Then
    Do While control <= num_termino
        temp = term_b
        term_b = term_a + term_b
        term_a = temp
        control = control + 1
    Loop
    MsgBox ("El término " & num_termino & " de la serie es: " & term_b)
Else
    Select Case num_termino
    Case 1:
        MsgBox ("El término " & num_termino & " de la serie es: 0")
    Case 2:
        MsgBox ("El término " & num_termino & " de la serie es: 1")
    Case 3:
        MsgBox ("El término solicitado no es válido")
    End Select
End If

End Sub
