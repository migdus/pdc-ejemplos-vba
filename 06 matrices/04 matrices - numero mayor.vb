' Autor: Miguel Dussán

' Escriba un programa que capture una matriz y determine el número mayor.

Option Explicit

Sub main()
    Dim m() As Integer ' Matriz
    
    Dim num_filas, num_columnas As Integer ' Almacenan las dimensiones de la matriz
    Dim mayor As Integer ' Numero mayor encontrado.
    Dim i, j As Integer ' Variables de control del ciclo
    Dim inicializa_mayor As Boolean
    inicializa_mayor = False
    
    num_filas = Val(InputBox("¿Cuántas filas desea introducir?"))
    
    num_columnas = Val(InputBox("¿Cuántas columnas desea introducir?"))
    
    'Definición de matriz
    'Se utilizan los valores del número de filas y de columnas capturados anteriormente para
    'declarar la matriz de forma dinámica.
    'Al valor de número de filas y columnas se les resta uno, pues los índices inician en cero.
    
    ReDim m(num_filas - 1, num_columnas - 1)

    ' Ciclo para capturar los valores de la matriz
    For i = 0 To num_filas - 1 ' Controla la posición de la fila
        For j = 0 To num_columnas - 1 ' Controla la posición de la columna
            m(i, j) = Val(InputBox("Escriba un valor para la posición " + CStr(i) + "," + CStr(j)))
        Next
    Next

    ' Ciclo para recorrer la matriz y encontrar el número mayor
     For i = 0 To num_filas - 1 ' Controla la posición de la fila
        For j = 0 To num_columnas - 1 ' Controla la posición de la columna
            
            ' Determina si se ha inicializado la variable mayor con el primer número de
            ' la matriz. Solo se hace una vez.
            If inicializa_mayor = False Then
                mayor = m(i, j)
                inicializa_mayor = True
            End If
            ' Determinar si el valor en m(i,j) es mayor que el valor almacenado
            ' en al variable mayor
            If m(i, j) > mayor Then
                mayor = m(i, j) ' Guardar como número mayor en caso verdadero
            End If
        Next
    Next

    MsgBox ("El número mayor es: " + CStr(mayor))

End Sub
