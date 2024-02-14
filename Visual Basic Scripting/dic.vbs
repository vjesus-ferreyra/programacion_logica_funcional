' Ejemplo 4: Creacion de Diccionarios
' Author: Víctor Jesús Torres Ferreyra
Dim Dic
' Declaración necesaria para crear objetos 
Set Dic = CreateObject("Scripting.Dictionary") ' Con esto creamos un diccionario
' Agragando elementos al diccionario
' RECORDAD: El diccionario es una estructura de datos de clave -> valor
Dic.Add "a", "Alicia"
Dic.Add "b", "Benito"
Dic.Add "c", "Carla"
Dic.Add "d", "Diego"
Dic.Add "e", "Estela"
Dic.Add "f", "Fernando"
Dic.Add "g", "Gustavo"
key = InputBox("Introduce una letra entre a - g") ' Lee la cadena y la guarda
MsgBox "Elemento seleccionado: " & Dic.Item(key) ' Imprime el elemento que coincida con la calve en el diccionario
