' Autor: Miguel Dussán

' Programa que busca un número dentro de un arreglo

Option Explicit

Sub main()
    Dim i, buscar, arreglo(1 To 10) As Integer
    
    ' Bandera para indicar si se encuentra el elemento
    Dim encontrado As Boolean
    encontrado = False
    
    ' Contenido del arreglo de ejemplo
    arreglo(1) = 123
    arreglo(2) = 45
    arreglo(3) = 1
    arreglo(4) = 23
    arreglo(5) = 5
    arreglo(6) = 8
    arreglo(7) = 11
    arreglo(8) = 13
    arreglo(9) = 87
    arreglo(10) = 2
    
    buscar = Val(InputBox("¿Qué número desea buscar?"))
    
    For i = 1 To 10
        ' Verificar si el arreglo en la posición i es igual al elemento buscado
        If arreglo(i) = buscar Then
            MsgBox ("Número encontrado en la posición " & i)
            encontrado = True ' marca el elemento como encontrado
        End If
    Next
    
    ' Mensaje que se muestra si no se encuentra el elemento
    If encontrado = False Then
        MsgBox ("El elemento solicitado no se encuentra en el arreglo")
    End If
    
End Sub
