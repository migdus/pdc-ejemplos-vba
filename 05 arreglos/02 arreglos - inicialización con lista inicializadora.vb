' Autor: Miguel Dussán

' Inicialización de un arreglo utilizando un ciclo

Option Explicit

Sub main()
    ' Declaración del arreglo de 10 elementos de tipo int
    Dim promedio, a(1 To 10) As Integer
    
    'Inicialización del arreglo de 10 elementos de tipo int
    
    a(1) = 56
    a(2) = 4
    a(3) = 2
    a(4) = 30
    a(5) = 12
    a(6) = 6
    a(7) = 4
    a(8) = 73
    a(9) = 49
    a(10) = 3
    
    promedio = 0
    
    Dim i As Integer ' Variable de control del ciclo
    
    Dim cadena As String
    
    cadena = ""
    
    For i = 1 To 10 Step 1
        promedio = promedio + a(i)
    Next
    
    promedio = promedio / 10
    
    MsgBox ("El promedio de los valores del arreglo es " & promedio)
    
End Sub