'Autor: Miguel Duss�n

'Este programa muestra el mensaje "Hola Mundo"
'Tambi�n muestra algunas de las caracter�sticas de este
'lenguaje de programaci�n.

'Esto es un comentario. Estas l�neas son omitidas cuando
'El compilador analiza y ejecuta el script.

Option Explicit ' Para forzar declaraci�n de variables

'Todos los procesos abren con la palabra Sub seguida del 
'nombre del proceso y luego unos par�ntesis que abren y cierran si
'no reciben par�metros. Por lo pronto no recibir�n par�metros.
' Los procesos terminan con las palabras End Sub.

Sub Main() ' Inicio del proceso de nombre Main

'La l�nea a continuaci�n muestra en pantalla la cadena que se
'encuentra dentro de las comillas dobles, utilizando el comando
'MsgBox().

MsgBox("�Hola Mundo!")

End Sub ' Fin del proceso