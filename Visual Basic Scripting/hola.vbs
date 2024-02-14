' Ejemplo 1: Hola mundo y entrada de datos por medio del usuario
' Author: Víctor Jesús Torres Ferreyra
' Hola amigos -> Formato para comentar una cosa
MsgBox("Hola amigos del Tec!") ' Creación de un cuadro de dialogo
REM Pide un nombre ' Comentario 
nombre = InputBox("Dame tu nombre") ' Creación de una variable y recibe por medio de un cuadro de dialogo una cadena 
MsgBox("Tu nombre es: " & nombre) ' Imprime el resultado y le concatena el valor de lo que tiene la variable 'nombre'
