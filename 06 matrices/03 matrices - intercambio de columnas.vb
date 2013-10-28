' Autor: Miguel Dussán

' Escriba un programa que realice el intercambio de valores de una matriz de
' la columna del centro con la primera columna. Si la matriz tiene un número
' par de columnas, asuma la columna del centor como aquella cuyo índice es la
' mitad del número de columnas.

Option Explicit

Sub main()
    Dim m() As Integer ' Matriz
    
    Dim num_filas, num_columnas As Integer ' Almacenan las dimensiones de la matriz
    
    Dim i, j, aux As Integer ' Variables de control del ciclo, y variable auxiliar
    
    Dim centro As Integer ' Variable para guardar la posición de la columna central
    
    num_filas = Val(InputBox("¿Cuántas filas desea introducir?"))
    
    num_columnas = Val(InputBox("¿Cuántas columnas desea introducir?"))
    
    'Definición de matriz
    'Se utilizan los valores del número de filas y de columnas capturados anteriormente para
    'declarar la matriz de forma dinámica.

    ReDim m(num_filas, num_columnas)

    ' Ciclo para capturar los valores de la matriz
    For i = 1 To num_filas ' Controla la posición de la fila
        For j = 1 To num_columnas ' Controla la posición de la columna
            m(i, j) = Val(InputBox("Escriba un valor para la posición " + CStr(i) + "," + CStr(j)))
        Next
    Next
    
    ' Operación de intercambio de valores de columnas
    
    centro = num_columnas \ 2 ' Backslash (\) es división entera, es diferente a /
    
    ' El programa debe recorrer fila por fila de la columna central, y luego intercambiar
    ' Esos valores con los de la primera columna, en la misma fila
    ' La primera columna tiene índice o posición 0
    
    For i = 1 To num_filas
        ' Intercambio de valores
        aux = m(i, 1) ' Almacena el valor de la fila i, columna 1 en una variable auxiliar
        m(i, 1) = m(i, centro) 'Asigna el valor de la columna central en esa fila a la fila de la primera columna
        m(i, centro) = aux 'Asignación del valor anterior de la fila i, columna 1 en la columna del centro en la misma fila
    Next
    
    
    ' Mostrar los valores en la hoja de cálculo
    ' Primer ciclo: recorre la matriz por filas
    For i = 1 To num_filas
        ' Segundo ciclo: recorre la matriz por columnas
        For j = 1 To num_columnas
            ' Escribir valores en la hoja de cálculo
            Cells(i, j) = m(i, j)
        Next
    Next
    
    ' Mostrar los valores en la hoja de cálculo
    ' Primer ciclo: recorre la matriz por filas
    For i = 1 To num_filas
        ' Segundo ciclo: recorre la matriz por columnas
        For j = 1 To num_columnas
            ' Escribir valores en la hoja de cálculo
            Cells(i, j) = m(i, j)
        Next
    Next

End Sub


