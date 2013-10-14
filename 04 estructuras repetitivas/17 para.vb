' Autor: Miguel Dussán

Option Explicit

Sub main()

Dim a As Integer
Dim acum As String

acum = "" ' Variable que va a acumular los resultados concatenados
    
    ' Inicia el ciclo For
    
    For a = 0 To 4 Step 2
        ' For es palabra clave
        ' a = 0 significa que el ciclo empieza a contar inicializando la variable a en cero
        ' To 4 señala que incrementará el valor de a hasta que llegue a 4, inclusive
        ' Step 2 indica el paso, que es de 2. Cuando el paso es de 1, esta parte de
                ' la instrucción puede omitirse.
        ' Variable cadena que acumula los valores del resultado
        acum = acum & a & " "
    Next
    ' Termina el ciclo For
    
    ' Imprime los valores concatenados en la variable acumuladora
    MsgBox (acum)

End Sub
