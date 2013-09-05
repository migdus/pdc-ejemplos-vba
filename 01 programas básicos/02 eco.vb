'Autor: Miguel Dussán

' Este programa pide una cadena de caracteres y luego la muestra en pantalla

Option Explicit ' Para forzar declaración de variables

Sub Main() ' Inicia función principal

Dim cadena As String ' Declaración de una cadena
cadena = "" ' Inicialización

' Muestra en una ventana la cadena "Escribe algo" y un
' campo para que el usuario ingrese algo. Además los
'botones aceptar y cancelar.
'La función InputBox captura la información de la caja
'de texto y esta la asigna a la variable que está a
'la izquierda, para este caso, cadena.

cadena = InputBox("Escribe algo")

'Mostramos la información que el usuario ingresó
'con la función MsgBox().

MsgBox (cadena)

End Sub ' Fin de la función principal
