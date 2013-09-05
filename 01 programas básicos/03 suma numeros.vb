'Autor: Miguel Dussán

' Este programa suma dos números

Option Explicit ' Para forzar declaración de variables

Sub Main() ' Inicia función principal

Dim a As Integer ' Declaración de entero (primer número)
Dim b As Integer ' Declaración de entero (segundo número)
Dim c As Integer ' Declaración de entero (almacenar el resultado)
Dim cadena_temp As String ' Declaración de cadena temporal, se utiliza para
													' guardar los datos capturados por teclado, y luego
													' convertirlos al tipo de dato Integer
													
a = 0 ' Inicialización de entero
b = 0 ' Inicialización de entero
c = 0 ' Inicialización de entero
cadena_temp = "" ' Inicialización de cadena con una cadena vacía

'Solicita el primer número y lo asigna inicialmente en la variable
'cadena_temp, y luego se convierte a Integer y se guarda en a
'esto se hace con ayuda de la función CInt() que toma como parámetro
' una cadena
cadena_temp = InputBox("Ingrese el primer número")
a = CInt(cadena_temp)

'Solicita el segundo número y lo asigna a la variable b, de la misma
'forma como se hizo con a
cadena_temp = InputBox("Ingrese el segundo número")
b = CInt(cadena_temp)

c = a + b ' Realiza la operación y la asigna a c

' Variable el resultado de la operación
Dim mensaje_resultado As String

' Las variables numéricas deben convertirse a cadenas
' Para eso se usa la función CStr() que recibe como parámetro
' La variable con el valor numérico, y lo convierte en cadena.
' Las cadenas se concatenan (unen) utilizando el operador "+".

mensaje_resultado = CStr(a) + " + " + CStr(b) + " = " + CStr(c)

'Se muestra el resultado
MsgBox (mensaje_resultado)

End Sub ' Fin de la función principal
