' Autor: Miguel Dussán

' Un vendedor ha hecho una serie de ventas y desea conocer la cantidad de ventas de $200 o
' menores, la cantidad de ventas que sean mayores a $200 e inferiores a $400, y las ventas de
' $400 o superiores. Diseñe un programa que provea al vendedor de esta información luego
' de leer los datos de entrada.

Option Explicit

Sub main()
    Dim venta, v_200, v_400, v_m400 As Integer
    
    venta = 1 ' Monto de una venta
    v_200 = 0 ' Cantidad de ventas de hasta $200
    v_400 = 0 ' Cantidad de ventas entre $200 y menores a $400
    v_m400 = 0 ' Cantidad de ventas entre $200 y menores a $400

    Dim cadena_resultado As String ' Cadena para almacenar el resultado

    Do While venta > 0
        venta = Val(InputBox("Ingrese el monto de una venta (0 o número negativo para terminar)"))
        If venta > 0 And venta <= 200 Then
            v_200 = v_200 + 1
        Else
            If venta > 0 And venta < 400 Then
                v_400 = v_400 + 1
            Else
                If venta > 0 And venta >= 400 Then
                    v_m400 = v_m400 + 1
                End If
            End If
        End If
    Loop
    
    
    cadena_resultado = "Ventas menores o iguales a $200: " & v_200 & vbNewLine
    cadena_resultado = cadena_resultado & "Ventas superiores a $200 y menores a $400: " & v_400 & vbNewLine
    cadena_resultado = cadena_resultado & "Ventas iguales o superiores a $400: " & v_m400
    
    MsgBox (cadena_resultado)

End Sub
