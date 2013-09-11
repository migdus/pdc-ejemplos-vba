'Autor: Miguel Dussán

' Este programa determina si un número entero es negativo
' utilizando una estructura Sí-Entonces

Option Explicit ' Para forzar declaración de variables

Sub Main() ' Inicio del proceso de nombre Main

	Dim numero As Integer ' Declaración de entero

	numero = 0 ' Inicialización

	'La línea a continuación muestra en pantalla la cadena que se
	'encuentra dentro de las comillas dobles, utilizando el comando
	' InputBox(). Además muestra un campo para que el usuario ingrese
	' un valor que será asignado a la variable numero.
	' La salida de InputBox se asigna a una cadena, y luego esta se convierte
	' a entero con la función Val 

	numero = Val(InputBox("Escriba un número"))

	' Estructura selectiva Sí-Entonces (If-Then).
	' El alcance de la estructura se identifica cierra con las
	' palabras clave: "End If"
	' La condición para que ejecute el código dentro de la estructura es
	' que el número sea menor que cero.

	If numero < 0 Then					' Abre la estructura
		MsgBox("El número es negativo")	' Imprime un mensaje si el número es negativo
	End If								' Cierra la estructura

End Sub ' Fin del proceso