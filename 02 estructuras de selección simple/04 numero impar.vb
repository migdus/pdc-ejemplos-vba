' Autor: Miguel Dussán

' Escriba un programa que determine si un número entero n es impar

Option Explicit ' Para forzar declaración de variables

Sub Main()
	Dim numero As Integer	' Declaración de entero

	numero = 0				' Inicialización
	
	' Impresión y captura de entrada en la variable numero
	numero = Val(InputBox("Escriba un número"))
	
	' Estructura selectiva con la condición:
	' si el número MOD 2 es diferente de cero, entonces el número es impar
	
	If numero mod 2 <> 0 Then
		MsgBox("El número es impar")
	End If
End Sub