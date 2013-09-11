' Autor: Miguel Duss�n

' Este programa suma dos n�meros

Option Explicit ' Para forzar declaraci�n de variables

Sub Main() ' Inicia funci�n principal

Dim a As Integer ' Declaraci�n de entero (primer n�mero)
Dim b As Integer ' Declaraci�n de entero (segundo n�mero)
Dim c As Integer ' Declaraci�n de entero (almacenar el resultado)

a = 0 ' Inicializaci�n de entero
b = 0 ' Inicializaci�n de entero
c = 0 ' Inicializaci�n de entero

' Solicita el primer n�mero y lo asigna a la variable a
' La funci�n InputBox("cadena") muestra un mensaje en pantalla con el texto
' cadena, y le permite al usuario ingresar un valor.
' La funci�n Val(.) convierte lo que est� como argumento a su respectivo
' valor num�rico. Se utiliza junto a InputBox como se muestra en la l�nea.

a = Val(InputBox("Ingrese el primer n�mero"))

'Solicita el segundo n�mero y lo asigna a la variable b

b = Val(InputBox("Ingrese el segundo n�mero"))

c = a + b ' Realiza la operaci�n y la asigna a c

' Variable el resultado de la operaci�n
Dim mensaje_resultado As String

' Las variables num�ricas se concatenan (se "juntan")
' utilizando el caracter ampersand (&) y esta informaci�n
' se almacena en una cadena que se mostrar� al final

mensaje_resultado = a & " + " & b & " = " & c

'Se muestra el resultado
MsgBox (mensaje_resultado)

End Sub ' Fin de la funci�n principal