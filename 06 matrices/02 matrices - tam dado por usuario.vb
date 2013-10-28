' Autor: Miguel Dussán

' Escriba un programa que permita crear una matriz del tamaño especificado por
' el usuario, y le permita introducir valores numéricos en ella.

Option Explicit

Sub main()

    Dim m() As Integer ' Matriz
    
    ' Almacenan las dimensiones de la matriz
    Dim num_filas, num_columnas As Integer
    
    Dim i, j As Integer ' Variables de control del ciclo
    
    num_filas = Val(InputBox("¿Cuántas filas desea introducir?"))
    
    num_columnas = Val(InputBox("¿Cuántas columnas desea introducir?"))
    
    'Definición de matriz
    'Se utilizan los valores del número de filas y de columnas capturados anteriormente para
    'declarar la matriz de forma dinámica.
    
    ReDim m(1 To num_filas, 1 To num_columnas)

    ' Ciclo para capturar los valores de la matriz
    For i = 1 To num_filas ' Controla la posición de la fila
        For j = 1 To num_columnas ' Controla la posición de la columna
            m(i, j) = Val(InputBox("Escriba un valor para la posición " + CStr(i) + "," + CStr(j)))
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

