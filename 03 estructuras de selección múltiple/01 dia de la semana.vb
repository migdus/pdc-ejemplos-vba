' Autor: Miguel Dussán

' Diseñe un algoritmo que dado el número ingresado por el usuario, muestre en
' pantalla el día de la semana al que corresponde (ejemplo: 3 es miércoles,
' 7 es domingo). Si el usuario ingresa una opción distinta a los números de la
' semana, deberá mostrar un mensaje de error.

Option Explicit ' Para forzar declaración de variables

Sub Main()

    Dim numero_dia As Integer ' Número de día
    
    ' Solicitar el número del día al usuario
    
    numero_dia = Val(InputBox("¿Número de día de la semana?"))
    
    Select Case numero_dia
        Case 1                          ' Si la variable tiene valor 1, imprime lunes
            MsgBox ("Lunes")
            
        Case 2                          ' Si la variable tiene valor 2, imprime martes
            MsgBox ("Martes")
            
        Case 3                          ' Si la variable tiene valor 3, imprime miércoles
            MsgBox ("Miércoles")
            
        Case 4                          ' Si la variable tiene valor 4, imprime jueves
            MsgBox ("Jueves")
            
        Case 5                          ' Si la variable tiene valor 5, imprime viernes
            MsgBox ("Viernes")
            
        Case 6                          ' Si la variable tiene valor 6, imprime sábado
            MsgBox ("Sábado")
            
        Case 7                          ' Si la variable tiene valor 7, imprime domingo
            MsgBox ("Domingo")
            
        Case Else ' Si no coincide ninguno de los valores anteriores con el contenido de la variable
            MsgBox ("No existe un día asociado a ese número")
            
    End Select

End Sub