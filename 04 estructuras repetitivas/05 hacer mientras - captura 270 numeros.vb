utor: Miguel Dussán

' Escriba un programa que para 270 números ingresados, calcule
' el promedio de los números pares y la suma de los números impares

Option Explicit

Sub main()

Dim captura As Integer ' Variable para guardar el dato escrito por el usuario

' Variables acumuladoras de números pares e impares
Dim acum_pares, acum_impares As Integer

' Variables que cuentan los números de pares
Dim cont_pares As Integer

' Variable de control de ciclo
Dim cont_capturados As Integer

' Inicialización de variables
acum_pares = 0
acum_impares = 0
cont_pares = 0
cont_capturados = 0

' Inicia ciclo mientras
' Se cuenta el número de veces que se repite el ciclo
' Va desde 0 hasta 269 para la variable cont_capturados

Do While cont_capturados < 270

    ' solicitud de dato a usuario
        captura = Val(InputBox("Escriba un número"))
	    
	        ' Si el número capturado es par
		    If captura Mod 2 = 0 Then
		        
			        ' Acumula el valor de números pares
				        acum_pares = acum_pares + captura
					        
						        'Cuenta cuántos números de los capturados son pares
							        ' Esto es indispensable para calcular el promedio al final
								        cont_pares = cont_pares + 1
									    Else
									            ' Si el número capturado es impar, acumula el valor de números impares
										            acum_impares = acum_impares + captura
											        End If

												    cont_capturados = cont_capturados + 1 ' Incrementa variable de control de ciclo
												    Loop ' Finaliza ciclo mientras

												    ' Imprimir resultados
												    ' Muestra el promedio como: acum_pares / cont_pares
												    ' La palabra vbNewLine agrega una nueva línea en el texto, de esta forma el resultado sale en dos renglones

												    MsgBox ("Promedio de números pares: " & acum_pares / cont_pares & vbNewLine & " Suma de impares: " & acum_impares)

												    End Sub
