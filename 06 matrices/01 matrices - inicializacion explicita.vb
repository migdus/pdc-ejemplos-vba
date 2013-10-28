' Autor: Miguel Dussán

' Escriba un programa que defina una matriz de 3x3 e inicialice sus valores en
' 2,3,4 para la primera fila, 6,7,4 para la segunda y 3,2,5 para la tercera.

Option Explicit

Sub main()

    'Definición de matriz:
    
    'Dim nombre_matriz(0 To numero_filas, 0 To numero_columnas) As tipo_dato
    'Aquí se define una matriz de 3 x 3
    Dim m(1 To 3, 1 To 3) As Integer
    
    Dim i, j As Integer ' Variables de control del ciclo
    
    ' Valores para la primera fila
    m(1, 1) = 2
    m(1, 2) = 3 ' Fila 2, Columna 2
    m(1, 3) = 4 
    
    ' Valores para la segunda fila
    m(2, 1) = 6
    m(2, 2) = 7
    m(2, 3) = 8 ' Fila 2, Columna 3
    
    ' Valores para la tercera fila
    m(3, 1) = 3 ' Fila 3, Columna 1
    m(3, 2) = 2
    m(3, 3) = 5
    
    ' Mostrar los valores en la hoja de cálculo
    ' Primer ciclo: recorre la matriz por filas
    For i = 1 To 3
        ' Segundo ciclo: recorre la matriz por columnas
        For j = 1 To 3
            ' Escribir valores en la hoja de cálculo
            Cells(i, j) = m(i, j)
        Next
    Next

End Sub
