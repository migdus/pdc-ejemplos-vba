'Autor: Miguel Dussán

'Este programa muestra el mensaje "Hola Mundo"
'También muestra algunas de las características de este
'lenguaje de programación.

'Esto es un comentario. Estas líneas son omitidas cuando
'El compilador analiza y ejecuta el script.

Option Explicit ' Para forzar declaración de variables

'Todos los procesos abren con la palabra Sub seguida del 
'nombre del proceso y luego unos paréntesis que abren y cierran si
'no reciben parámetros. Por lo pronto no recibirán parámetros.
' Los procesos terminan con las palabras End Sub.

Sub Main() ' Inicio del proceso de nombre Main

'La línea a continuación muestra en pantalla la cadena que se
'encuentra dentro de las comillas dobles, utilizando el comando
'MsgBox().

MsgBox("¡Hola Mundo!")

End Sub ' Fin del proceso