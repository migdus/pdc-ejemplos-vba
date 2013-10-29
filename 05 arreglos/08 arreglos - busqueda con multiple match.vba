' Autor: Miguel Dussán

' Escriba un programa que capture números en un arreglo. Después el programa
' debe pedir un número que debe buscar en el arreglo y retornar un arreglo donde
' marque las posiciones en las que este se encuentra.

Option Explicit

Sub main()
' arreglo : el arreglo que almacenará los elementos que el usuario digite
' tam : cantidad de elementos que el usuario ingresará
' buscar : término de búsqueda
' i : variable de control del ciclo
' encontre_uno : bandera que marca si se encontró un elemento
' encontrado : arreglo de banderas que marca las posiciones en
' las que está el elemento de búsqueda
Dim arreglo(), tam, buscar, i As Integer
Dim encontre_uno, encontrado() As Boolean

encontre_uno = False

tam = Val(InputBox("¿Cantidad de números a ingresar?"))

ReDim arreglo(tam)
ReDim encontrado(tam)

For i = 1 To tam
    ' Captura de valores
    arreglo(i) = Val(InputBox("Escriba un número para la posición " & i))
    encontrado(i) = False ' inicialización de arreglo de banderas
Next

' Captura del término de búsqueda
buscar = Val(InputBox("¿Qué número quiere buscar?"))

For i = 1 To tam
    ' Si el elemento en el arreglo en la posición i coincide
    ' con el elemento de búsqueda...
    If arreglo(i) = buscar Then
        encontrado(i) = True ' marcar en la matriz de encontrados en esa posición
        encontre_uno = True ' marcar que se encontró al menos un elemento
    End If
Next

Dim cadena As String ' cadena para almacenar los resultados
cadena = ""

' Si se encontró al menos un elemento
If encontre_uno = True Then
    For i = 1 To tam
        ' Revisar en la matriz donde se marcan las posiciones si está marcado
        ' en una posición específica como verdadero
        If encontrado(i) = True Then
            ' concatenar resultados
            cadena = cadena & " " & i
        End If
    Next
    ' Mostrar resultados
    MsgBox ("Elementos encontrados en las posiciones " & cadena)
Else
    MsgBox ("No se encontraron elementos")
End If
End Sub