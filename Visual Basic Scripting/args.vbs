' Ejemplo 2: Recibimeinto de numero de parametros
' Author: Víctor Jesús Torres Ferreyra
' Éjemplo de numero de argumentos
Dim Args, ArgList ' Asignación de una variable y/o argumento
Set Args = WScript.Arguments ' Asignamos el valor a una variable
For i=0 to Args.Count - 1 ' Creación de un ciclo for
	' ¿Qué hace este ciclo?
	' Lo que hace es que el valor que tenga 'ArgList' se concatena por el valor que tenga i,
	' luego este se concatena ":",
	' luego se concatena el valor correspondiente a su posicion,
	' al final se le agrega el retorno de carro 'chr(13)'
	ArgList = ArgList & i & ": " & Args(i) & chr(13) 
Next ' indica el fin de un ciclo for
' Imprimir un menjaje concatenando el retorno de carro y el valor de la longitud de Args
MsgBox "Numero de argumentos: " & Args.Count & chr(13) & ArgList
