' Ejemplo 3: Creación de un archivo
' Author: Víctor Jesús Torres Ferreyra
Dim MSword, WSHSell ' Creación de variables 
Set WSHShell = WScript.CreateObject("WScript.Shell") ' Creación y asignacion de un objeto de tipo Shell
Set MSWord = WSCript.CreateObject("Word.Basic") ' Creación y asignación de un objeto de tipo 'Word'
MSWord.FileNew("Normal") ' Creamos un archivo basado en la plantilla 'Normal'
' Insertamos cadenas a este documento, en este caso dos cadenas separadas por el retorno de carro
MSWord.Insert("Programacion Logica y Funcional" & Chr(13))
MSWord.Insert("Profesor: M.I. Jorge Eduardo Carrion")
' Guardamos este archivo en un directorio en especifico (SendTo en este caso) y con un nombre y formato en especifico
MSWord.FileSaveAs(WSHShell.SpecialFolders(12) & "\lp.docx")
MSWord.FileClose ' Cerramos el buffer del archivo
