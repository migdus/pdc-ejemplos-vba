' Autor: Miguel Dussán

' Escriba un programa que capture 3 números y los muestre de forma ascendente

Option Explicit ' Para forzar declaración de variables

Sub Main()

    ' Declaración de las tres variables de entrada
    ' Si las variables son del mismo tipo, pueden declararse en la misma línea
    ' De forma alternativa, pueden declararse cada una en su línea
    
    Dim a, b, c As Integer

    ' Impresión y captura de entrada en la variable a
    a = Val(InputBox("Escriba un número"))
    
    ' Impresión y captura de entrada en la variable b
    b = Val(InputBox("Escriba un número"))
    
    ' Impresión y captura de entrada en la variable c
    c = Val(InputBox("Escriba un número"))
    
    If a < b And b < c Then ' Primer si entonces
        MsgBox (a & " " & b & " " & c)
    Else ' Si no del primer si entonces
        If a < c And c < b Then ' Segundo si entonces, anidado
            MsgBox (a & " " & c & " " & b)
        Else  '  Si no del segundo si entonces
            If b < a And a < c Then ' Tercer si entonces, anidado
                MsgBox (b & " " & a & " " & c)
            Else  ' Si no del tercer si entonces
                If b < c And c < a Then ' Cuarto si entonces, anidado
                    MsgBox (b & " " & c & " " & a)
                Else  ' Si no del cuarto si entonces
                    If c < a And a < b Then ' Quinto Tercer si entonces, anidado
                        MsgBox (c & " " & a & " " & b)
                    Else  ' Si no del quinto si entonces
                        If c < b And b < a Then ' Sexto Tercer si entonces, anidado
                            MsgBox (c & " " & b & " " & a)
                        Else  ' Si no del sexto si entonces
                            MsgBox ("Existe un error con los números ingresados. ¿Existen dos números iguales?")
                        End If ' Cierre del sexto si entonces
                    End If ' Cierre del quinto si entonces
                End If  ' Cierre del cuarto si entonces
            End If  ' Cierre del tercer si entonces
        End If  ' Cierre del segundo si entonces
    End If  ' Cierre del primer si entonces
    
End Sub
