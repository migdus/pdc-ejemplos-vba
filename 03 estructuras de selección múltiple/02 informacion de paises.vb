' Autor: Miguel Dussán

' Diseñe un algoritmo que dado el número ingresado por el usuario, muestre en
' pantalla el día de la semana al que corresponde (ejemplo: 3 es miércoles,
' 7 es domingo). Si el usuario ingresa una opción distinta a los números de la
' semana, deberá mostrar un mensaje de error.

Option Explicit ' Para forzar declaración de variables

Sub Main()

    Dim seleccion As Integer  ' Variable que guarda la selección del usuario
    
    Dim cadena As String ' Cadena para almacenar la información que se mostrará al usuario
    
    cadena = "Seleccione el número del país para conocer su capital: 1. Colombia"
    cadena = cadena & "2. Venezuela 3. Perú 4. Brasil"
    
    ' Solicitar la opción al usuario
    
    seleccion = Val(InputBox(cadena))
    
    Select Case seleccion
        Case 1
            MsgBox ("La capital de Colombia es Bogotá")
            
        Case 2
            MsgBox ("La capital de Venezuela es Caracas")
            
        Case 3
            MsgBox ("La capital de Perú es Lima")
            
        Case 4
            MsgBox ("La capital de Brasil es Brasilia")

        ' Si no coincide ninguno de los valores anteriores con el contenido de la variable

        Case Else
            MsgBox ("La opcion seleccionada no corresponde a una de las ofrecidas")
            
    End Select

End Sub
