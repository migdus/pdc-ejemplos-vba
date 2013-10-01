' Autor: Miguel Dussan

' El Ministerio de Agricultura colombiano tiene los datos resumidos en miles de
' toneladas de cosecha recolectada en los pisos términos cálido, templado y frío
' del país para los trimestres del año anterior. Diseñe un programa que utilizando
' estructuras repetitivas permita al ministerio calcular:
'
' a) El promedio trimestral de cosecha recolectada para el piso térmico templado.
' b) El mejor y peor trimestre en cantidad de cosecha recolectada para todo el país.
' c) El piso término que tuvo la mayor cantidad de toneladas recolectadas en el último trimestre del año.
' d) (Ejercicio) El piso término que tuvo la menor cantidad de toneladas recolectadas en el segundo trimestre del año.
' e) (Ejercicio) El promedio del tercer trimestre de cosecha recolectada para el piso térmico templado.

Option Explicit

Sub main()
    
' Toneladas recolectadas en un trimestre dado
Dim calido, templado, frio As Integer

Dim prom_frio, control, total_trim  As Integer
control = 0 ' Variable de control de ciclo
prom_frio = 0 ' Promedio de recolección para el clima frío

total_trim = 0 ' Total recolectado en un trimestre

' Guarda el número del mejor y del peor trimestre.
' El valor inicial es 1, asumiendo que el primero será el peor. Si existe
' un trimestre peor que este, el programa lo modificará

Dim mejor_trim, peor_trim As Integer
mejor_trim = 1
peor_trim = 1

' Cantidad recolectada para el mejor y peor trimestre
Dim mejor_trim_cant_recol, peor_trim_cant_recol As Integer
mejor_trim_cant_recol = 0
peor_trim_cant_recol = 0

' Esta bandera permite detectar si es el primer bucle que se está ejecutando. Si es verdadero,
' entonces se inicializan las variables de cantidad recolectada para el mejor y peor trimestre.
' Este bloque se ejecuta una sola vez
Dim band_primera_iter As Boolean
band_primera_iter = True

' Indica el piso térmico que tuvo la mayor recolección en el último trimestre
' 1 para cálido, 2 para templado y 3 para frío
Dim mayor_recol_ultimo_trim As Integer
mayor_recol_ultimo_trim = 0

' El año tiene 4 trimestres
Do While control < 4

        ' Solicitud de toneladas recolectadas en un trimestre
        calido = Val(InputBox("Toneladas recolectadas para el clima calido para el trimestre " & control + 1))
        templado = Val(InputBox("Toneladas recolectadas para el clima templado para el trimestre " & control + 1))
        frio = Val(InputBox("Toneladas recolectadas para el clima frio para el trimestre " & control + 1))
        
        prom_frio = prom_frio + frio
        
        total_trim = calido + templado + frio
        
        If band_primera_iter = True Then
            mejor_trim_cant_recol = total_trim
            peor_trim_cant_recol = total_trim
            band_primera_iter = False
        Else
            If total_trim > mejor_trim_cant_recol Then
                mejor_trim = control + 1
                mejor_trim_cant_recol = total_trim
            Else
                If total_trim < peor_trim_cant_recol Then
                    peor_trim = control + 1
                    peor_trim_cant_recol = total_trim
                End If
            End If
        End If
        
        If control = 3 Then
            If calido > templado And calido > frio Then
                mayor_recol_ultimo_trim = 1
            Else
                If templado > calido And templado > frio Then
                    mayor_recol_ultimo_trim = 2
                Else
                    If frio > calido And frio > templado Then
                        mayor_recol_ultimo_trim = 3
                    End If
                End If
            End If
        End If
        control = control + 1
Loop

Dim cadena_resultado As String

cadena_resultado = "Resultados" & vbNewLine
cadena_resultado = cadena_resultado & " El promedio trimestral de cosecha recolectada para el piso térmico templado es: "
cadena_resultado = cadena_resultado & prom_frio / 4 & vbNewLine

cadena_resultado = cadena_resultado & "El mejor trimestre fue el numero: " & mejor_trim
cadena_resultado = cadena_resultado & " con un total de " & mejor_trim_cant_recol & " toneladas." & vbNewLine

cadena_resultado = cadena_resultado & "El peor trimestre fue el numero: " & peor_trim & " con un total de "
cadena_resultado = cadena_resultado & peor_trim_cant_recol & " toneladas." & vbNewLine

cadena_resultado = cadena_resultado & "El piso termico con mayor recoleccion en el ultimo trimestre fue: "
Select Case mayor_recol_ultimo_trim
    Case 1:
        cadena_resultado = cadena_resultado & "calido"
    Case 2:
        cadena_resultado = cadena_resultado & "templado"
    Case 3:
        cadena_resultado = cadena_resultado & "frio"
End Select

MsgBox (cadena_resultado)

End Sub

