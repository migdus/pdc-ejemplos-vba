' Autor: Miguel Dussán

' Una compañía de seguridad ofrece guardaespaldas para altos perfiles (empresarios, políticos, estrellas
' de farándula y cualquier otra persona que pueda permitirse pagar estos servicios). Cada vez que un
' guardaespaldas asiste a uno de estos clientes, el cliente firma una planilla donde se registra el número
' de horas y la jornada en la que el guardaespaldas lo estuvo acompañando. Al final del mes, el guardaespaldas
' entrega la planilla a la empresa para que esta le pague las horas trabajadas, discriminadas en las siguientes
' jornadas: diurno, nocturno y festivo, cada una con una tarifa diferente por hora, como se muestra al final.
' Diseñe un programa que permita ingresar cuántas horas trabajó por jornada cada guardaespaldas, y calcule:
' cuántas horas trabajó el empleado en ese mes, cuál es el total que se le debe pagar al empleado, qué porcentaje
' de lo que se le debe pagar corresponde a cada jornada, y cuánto dinero pagarán los clientes a la empresa de
' seguridad en ese mes.

' Valores a pagar al guarda espaldas por jornada
' ==================================
' diurno              $10.000/hora
' nocturno          $15.000/hora
' festivo             $20.000/hora

' Valor a cobrar al cliente
' ==================
' diurno                $30.000/hora
' nocturno            $50.000/hora
' festivo               $70.000/hora

' Ejercicios sugeridos

' 1. ¿Cuál es el total que recibe la empresa después de pagarle a todos sus empleados?
' 2. ¿Cuál de las jornadas generó mayores ingresos para la empresa?


Option Explicit

Sub main()

' cantidad: número de empleados
' control: variable de control del ciclo
'diurno, nocturno, festivo: horas trabajadas para un empleado dado
' total: total de horas trabajadas para un trabajador
' pagar: total a pagar a un empleado
' recibido: acumuladora, dinero cobrado a los clientes

    Dim cantidad, control, diurno, nocturno, festivo, total, pagar, recibido As Double
    
    Dim cadena_resultado As String
    
    control = 0
    recibido = 0
    
    cantidad = Val(InputBox("Digite la cantidad de empleados cuyas horas quiere registrar"))
    
    Do While control < cantidad
      
        diurno = Val(InputBox("Registro de horas diurnas trabajadas para el empleado " & control + 1))
        nocturno = Val(InputBox("Registro de horas nocturnas trabajadas para el empleado " & control + 1))
        festivo = Val(InputBox("Registro de horas en día festivo trabajadas para el empleado " & control + 1))
        
        total = diurno + nocturno + festivo
        
        ' total a pagar, de acuerdo a las condiciones del problema
        pagar = (diurno * 10 + nocturno * 15 + festivo * 20) * 1000
        
        cadena_resultado = ""
        cadena_resultado = "Resultados para el empleado No. " & control + 1 & vbNewLine
        cadena_resultado = cadena_resultado & "Horas trabajadas en el mes: " & total & vbNewLine
        cadena_resultado = cadena_resultado & "Se le debe pagar: $" & pagar & vbNewLine
        
        ' Cálculo de porcentaje para cada jornada.
        ' Se calcula con base en el tiempo trabajado por jornada y el total de horas trabajadas en el mes
        
        cadena_resultado = cadena_resultado & "Porcentaje correspondiente a jornada diurna: " & diurno * 100 / total & "%" & vbNewLine
        cadena_resultado = cadena_resultado & "Porcentaje correspondiente a jornada nocturna: " & nocturno * 100 / total & "%" & vbNewLine
        cadena_resultado = cadena_resultado & "Porcentaje correspondiente a jornada de festivos: " & festivo * 100 / total & "%" & vbNewLine
        
        MsgBox (cadena_resultado)
        
        recibido = recibido + (diurno * 30 + nocturno * 50 + festivo * 70) * 1000
        
        control = control + 1
    Loop
    
    MsgBox ("El total pagado por los clientes a la empresa fue de: $" & recibido)
End Sub