' Autor: Miguel Dussán

Option Explicit

Sub main()
    Dim opcion As Integer
    Dim temp As String
    Dim texto_menu As String
    
    opcion = 0 'Variable para almacenar la opción escogida por el usuario
    
    texto_menu = "MENÚ DEL DÍA" & vbNewLine & vbNewLine
    texto_menu = texto_menu & "Seleccione el número de la opción que desea:" & vbNewLine & vbNewLine
    texto_menu = texto_menu & "1. ARROZ A LA GIRASOL" & vbNewLine & vbNewLine
    texto_menu = texto_menu & "Un delicioso plato de arroz con un huevo frito encima." & vbNewLine & vbNewLine
    texto_menu = texto_menu & "2. ATÚN CON PASTA" & vbNewLine & vbNewLine
    texto_menu = texto_menu & "El tradicional plato de soltero, con tallarines." & vbNewLine & vbNewLine
    texto_menu = texto_menu & "3. CALENTADO" & vbNewLine & vbNewLine
    texto_menu = texto_menu & "El almuerzo de ayer, hoy"
    
    'Inicia ciclo do-while
    Do
        'Se ejecuta todo lo que se encuentra en este bloque al menos una vez.
        'Luego verifica la condición, y si esta es verdadera, vuelve y ejecuta
        'el bloque de código dentro del do-while.
        'Si en algún momento la condición es falsa, la ejecución continúa por
        'fuera del ciclo do-while.
        
        opcion = Val(InputBox(texto_menu))

        'Mientras el usuario escriba una opción que no se encuentre dentro del rango
    Loop While opcion < 1 Or opcion > 3
    'Termina ciclo do-while
    
    MsgBox ("Ha seleccionado la opción " & opcion)
End Sub
