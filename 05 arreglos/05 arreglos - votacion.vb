' Autor Miguel Dussan

' Escriba un programa que para un número determinado de votantes ingresados
' por el usuario, permita capturar los votos para 5 candidatos y los
' almacene en un arreglo. Con esta información el programa debe calcular
' el ganador de la elección, teniendo en cuenta que gana el candidato que
' haya obtenido más del 50% de los votos.

Option Explicit

Sub main()

    Dim voto, i As Integer
    Dim votantes, votos(1 To 5), resultados(1 To 5) As Double
    Dim existe_ganador As Boolean
    existe_ganador = False ' Indica si se encuentra un ganador
    
    ' Inicialización de los arreglos
    For i = 1 To 5
        votos(i) = 0
        resultados(i) = 0
    Next
    
    votantes = Val(InputBox("¿Cuántas personas desean votar?"))
    
    For i = 1 To votantes
        Do
            voto = Val(InputBox("Ingrese el voto para el ciudadano " & i))
        Loop While voto < 1 Or voto > 5
        
        votos(voto) = votos(voto) + 1 ' Se suma uno al candidato seleccionado
    Next
    
    ' Calcular los resultados
    For i = 1 To 5
        resultados(i) = votos(i) * 100 / votantes
    Next
    
    ' Verificar quién sacó más del 50%
    For i = 1 To 5
        If resultados(i) > 50 Then
            MsgBox ("El ganador fue el candidato " & i & " con un total de " & resultados(i) & "%" & " y un total de " & votos(i))
            existe_ganador = True
        End If
    Next
    
    If existe_ganador = False Then
        MsgBox ("No existe ganador")
    End If

End Sub