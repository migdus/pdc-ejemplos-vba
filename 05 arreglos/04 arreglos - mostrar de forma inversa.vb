' Autor: Miguel Dussán

Option Explicit

Sub main()
    Dim promedio As Double
    
    ' a() es un arreglo sin tamaño definido (aún)
    Dim i, tam, a() As Integer
    Dim cadena As String
    
    tam = Val(InputBox("¿Cuántos números desea ingresar?"))
    
    ReDim a(tam) ' Fijar el tamaño de acuerdo a lo especificado por el usuario
    
    ' Captura de los valores
    
    For i = 1 To tam
        a(i) = Val(InputBox("Ingrese un número"))
    Next
    
    For i = tam To 1 Step -1
        cadena = cadena & a(i) & " "
    Next

    MsgBox (cadena)
End Sub