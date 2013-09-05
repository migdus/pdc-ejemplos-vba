'Autor: Miguel Duss�n

' Este programa suma dos n�meros

Option Explicit ' Para forzar declaraci�n de variables

Sub Main() ' Inicia funci�n principal

Dim a As Integer ' Declaraci�n de entero (primer n�mero)
Dim b As Integer ' Declaraci�n de entero (segundo n�mero)
Dim c As Integer ' Declaraci�n de entero (almacenar el resultado)
Dim cadena_temp As String ' Declaraci�n de cadena temporal, se utiliza para
													' guardar los datos capturados por teclado, y luego
													' convertirlos al tipo de dato Integer
													
a = 0 ' Inicializaci�n de entero
b = 0 ' Inicializaci�n de entero
c = 0 ' Inicializaci�n de entero
cadena_temp = "" ' Inicializaci�n de cadena con una cadena vac�a

'Solicita el primer n�mero y lo asigna inicialmente en la variable
'cadena_temp, y luego se convierte a Integer y se guarda en a
'esto se hace con ayuda de la funci�n CInt() que toma como par�metro
' una cadena
cadena_temp = InputBox("Ingrese el primer n�mero")
a = CInt(cadena_temp)

'Solicita el segundo n�mero y lo asigna a la variable b, de la misma
'forma como se hizo con a
cadena_temp = InputBox("Ingrese el segundo n�mero")
b = CInt(cadena_temp)

c = a + b ' Realiza la operaci�n y la asigna a c

' Variable el resultado de la operaci�n
Dim mensaje_resultado As String

' Las variables num�ricas deben convertirse a cadenas
' Para eso se usa la funci�n CStr() que recibe como par�metro
' La variable con el valor num�rico, y lo convierte en cadena.
' Las cadenas se concatenan (unen) utilizando el operador "+".

mensaje_resultado = CStr(a) + " + " + CStr(b) + " = " + CStr(c)

'Se muestra el resultado
MsgBox (mensaje_resultado)

End Sub ' Fin de la funci�n principal
