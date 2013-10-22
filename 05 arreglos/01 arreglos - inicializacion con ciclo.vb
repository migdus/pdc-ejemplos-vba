' Autor: Miguel Dussán

' Inicialización de un arreglo utilizando un ciclo

Option Explicit

Sub main()
    ' Declaración del arreglo de 10 elementos de tipo int
    Dim a(1 To 10) As Integer
    
    Dim i As Integer ' Variable de control del ciclo
    
    Dim cadena As String
    
    cadena = ""
    
    ' Inicialización del arreglo
    ' Se debe tomar cada una una de las posiciones y asignarles un valor
    
    For i = 1 To 10 Step 1
        a(i) = 0  ' Valor inicial para la posición i-ésima del arreglo
    Next
    
    ' Mostrar la información del arreglo en un MsgBox
    
    For i = 1 To 10
        cadena = cadena & a(i) & " "
    Next
        
    MsgBox (cadena)
    
End Sub
