'Autor: Miguel Duss�n

' Este programa pide una cadena de caracteres y luego la muestra en pantalla

Option Explicit ' Para forzar declaraci�n de variables

Sub Main() ' Inicia funci�n principal

Dim cadena As String ' Declaraci�n de una cadena
cadena = "" ' Inicializaci�n

' Muestra en una ventana la cadena "Escribe algo" y un
' campo para que el usuario ingrese algo. Adem�s los
'botones aceptar y cancelar.
'La funci�n InputBox captura la informaci�n de la caja
'de texto y esta la asigna a la variable que est� a
'la izquierda, para este caso, cadena.

cadena = InputBox("Escribe algo")

'Mostramos la informaci�n que el usuario ingres�
'con la funci�n MsgBox().

MsgBox (cadena)

End Sub ' Fin de la funci�n principal
