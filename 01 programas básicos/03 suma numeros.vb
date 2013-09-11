' Autor: Miguel Dussán

' Este programa suma dos números

Option Explicit ' Para forzar declaración de variables

Sub Main() ' Inicia función principal

Dim a As Integer ' Declaración de entero (primer número)
Dim b As Integer ' Declaración de entero (segundo número)
Dim c As Integer ' Declaración de entero (almacenar el resultado)

a = 0 ' Inicialización de entero
b = 0 ' Inicialización de entero
c = 0 ' Inicialización de entero

' Solicita el primer número y lo asigna a la variable a
' La función InputBox("cadena") muestra un mensaje en pantalla con el texto
' cadena, y le permite al usuario ingresar un valor.
' La función Val(.) convierte lo que está como argumento a su respectivo
' valor numérico. Se utiliza junto a InputBox como se muestra en la línea.

a = Val(InputBox("Ingrese el primer número"))

'Solicita el segundo número y lo asigna a la variable b

b = Val(InputBox("Ingrese el segundo número"))

c = a + b ' Realiza la operación y la asigna a c

' Variable el resultado de la operación
Dim mensaje_resultado As String

' Las variables numéricas se concatenan (se "juntan")
' utilizando el caracter ampersand (&) y esta información
' se almacena en una cadena que se mostrará al final

mensaje_resultado = a & " + " & b & " = " & c

'Se muestra el resultado
MsgBox (mensaje_resultado)

End Sub ' Fin de la función principal