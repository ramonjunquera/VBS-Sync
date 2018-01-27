'Título: Sincronización de carpetas
'Autor: Ramón Junquera
'Versión: 20180127
'Descripción:
'  Sincronización de dos carpetas en modo espejo (folder1 -> folder2)
'  teniendo en cuenta el tamaño y la fecha

'Definición de carpetas
sourceFolder     ="D:\borrame\folder1\"
destinationFolder="D:\borrame\folder2\"
logFolder        ="D:\borrame\"

'Definición de variables globales
'Objeto de gestión de archivos
Set fso = CreateObject("Scripting.FileSystemObject")
'Archivo de log
Dim logFile
'Nombre del archivo de log
Dim logFileName
'Número de errores
dim errorCount


'Definición de clases
Class fileStruct
  'Estructura de archivo
  public size
  public date
End Class


'Definición de funciones
function createFolder(path)
  'Crea una carpeta
  line=now & " creando carpeta " & path & " ... "
  'Si hay errores...continuaremos
  On Error Resume Next
  'Creamos la carpeta
  fso.CreateFolder(path)
  'Si no hubo errores...
  if Err.Number = 0 Then
    line=line & "ok"
  else 'Si hubo errores...
    'Componemos la línea a escribir en el log
    line=line & "ERROR : " & Err.Description
    'Limpiamos el flag de errores
    Err.Clear
    'Aumentamos el número de errores
    errorCount=errorCount+1
  end if
  'Escribimos la línea en el archivo de log
  logFile.Write line & vbCrLf
end Function

function deleteFolder(path)
  'Borra una carpeta con todo su contenido
  line=now & " borrando carpeta " & path & " ... "
  'Si hay errores...continuaremos
  On Error Resume Next
  'Borramos la carpeta
  fso.DeleteFolder(path),true
  'Si no hubo errores...
  if Err.Number = 0 Then
    line=line & "ok"
  else 'Si hubo errores...
    'Componemos la línea a escribir en el log
    line=line & "ERROR : " & Err.Description
    'Limpiamos el flag de errores
    Err.Clear
    'Aumentamos el número de errores
    errorCount=errorCount+1
  end if
  'Escribimos la línea en el archivo de log
  logFile.Write line & vbCrLf
end Function

function copyFile(path1,path2)
  'Copia un archivo
  line=now & " copiando archivo " & path1 & " ... "
  'Si hay errores...continuaremos
  On Error Resume Next
  'Copiamos el archivo
  fso.CopyFile path1,path2,true
  'Si no hubo errores...
  if Err.Number = 0 Then
    line=line & "ok"
  else 'Si hubo errores...
    'Componemos la línea a escribir en el log
    line=line & "ERROR : " & Err.Description
    'Limpiamos el flag de errores
    Err.Clear
    'Aumentamos el número de errores
    errorCount=errorCount+1
  end if
  'Escribimos la línea en el archivo de log
  logFile.Write line & vbCrLf
end Function

function deleteFile(path)
  'Copia un archivo
  line=now & " borrando archivo " & path & " ... "
  'Si hay errores...continuaremos
  On Error Resume Next
  'Borramos el archivo
  fso.DeleteFile(path),true
  'Si no hubo errores...
  if Err.Number = 0 Then
    line=line & "ok"
  else 'Si hubo errores...
    'Componemos la línea a escribir en el log
    line=line & "ERROR : " & Err.Description
    'Limpiamos el flag de errores
    Err.Clear
    'Aumentamos el número de errores
    errorCount=errorCount+1
  end if
  'Escribimos la línea en el archivo de log
  logFile.Write line & vbCrLf
end Function

function copyFolder(path1,path2)
  'Copia una carpeta
  line=now & " copiando carpeta " & path1 & " ... "
  'Si hay errores...continuaremos
  On Error Resume Next
  'Copiamos el archivo
  fso.CopyFolder path1,path2,true
  'Si no hubo errores...
  if Err.Number = 0 Then
    line=line & "ok"
  else 'Si hubo errores...
    'Componemos la línea a escribir en el log
    line=line & "ERROR : " & Err.Description
    'Limpiamos el flag de errores
    Err.Clear
    'Aumentamos el número de errores
    errorCount=errorCount+1
  end if
  'Escribimos la línea en el archivo de log
  logFile.Write line & vbCrLf
end Function

Function syncFolder(path1,path2)
  'Sincroniza el contenido de una carpeta en otra
  'Ej: syncFolder("D:\Ra\sourceFolder\","D:\Ra\destinationFolder\")
  'Si el path1 existe...
  if fso.FolderExists(path1) then
    'Si el path2 no existe
    if(not fso.FolderExists(path2)) then
      '...lo creamos
      createFolder(path2)
    end if
    'Creamos un diccionario para guardar los archivos de la carpeta destino
    Set dic2=CreateObject("Scripting.Dictionary")
    'Abrimos la carpeta destino
    Set folder2=fso.GetFolder(path2)
    'Recorremos todos los archivos de la carpeta destino
    For Each file In folder2.Files
      'Creamos un nuevo nodo para guardar la información de un archivo
      Set node = new fileStruct
      'Añadimos la información del archivo al nodo
      node.size=file.Size
      node.date=file.DateLastModified
      'Añadimos el archivo nodo al diccionario usando como clave el nombre del archivo
      dic2.Add file.Name,node
    Next 'file
    'Recorremos las subcarpetas de la carpeta destino
    For Each folder In folder2.SubFolders
      'Creamos un nuevo nodo para guardar la información de una carpeta
      Set node = new fileStruct
      'Añadimos la información de la carpeta al nodo
      node.size=-1 'Las carpetas no tienen tamaño
      node.date=0 'No tenemos en cuenta la fecha en las carpetas
      'Añadimos la carpeta al diccionario usando como clave el nombre de la carpeta
      dic2.Add folder.Name,node
    Next 'folder
    'Abrimos la carpeta origen
    Set folder1=fso.GetFolder(path1)
    'Recorremos todos los archivos de la carpeta origen
    For Each file in folder1.Files
      'Si en destino ya existe un archivo o carpeta con el mismo nombre...
      if(dic2.Exists(file.Name)) Then
        '..si el destino es una carpeta...
        if(dic2.Item(file.Name).size=-1) Then
          '...tenemos que borrar esa carpeta, porque si no, no nos dejará copiar el archivo
          deleteFolder(path2 & file.Name)
        end if
        'Tenemos que decidir si copiamos el archivo
        'Si el tamaño o la fecha no coinciden...
        if (file.Size <> dic2.item(file.Name).size) or (file.DateLastModified <> dic2.Item(file.Name).date) Then
          '...copiamos el archivo sobreescribiendolo
          copyFile path1 & file.Name,path2  
        end if
        'Eliminamos la referencia del diccionario para no procesarla después
        dic2.Remove(file.Name)
      else 'No existe el archivo en destino...
        '...lo copiamos
        copyFile path1 & file.Name,path2
      end if
    Next 'file
    'Los nodos de archivo que contiene el diccionario corresponden a archivo inexistentes
    'en origen. Debemos borrarlos en destino.
    'Recorremos todos los nodos del diccionario
    For Each key In dic2.Keys
      'Si el nodo es de un archivo...
      if(dic2.Item(key).size<>-1) Then
        '...eliminamos el archivo
        deleteFile(path2 & key)
        '...y el nodo del diccionario
        dic2.Remove(key)
      end if
    Next 'key
    'En el diccionario sólo quedan nodos de carpetas
    'Recorremos las carpetas de origen
    For Each folder in folder1.SubFolders
      'Si la carpeta existe en destino...
      if dic2.Exists(folder.Name) Then
        '...eliminamos la referencia del diccionario
        dic2.Remove(folder.Name)
        'Sincronizamos las carpetas
        syncFolder path1 & folder.Name & "\",path2 & folder.Name & "\"
      else 'La carpeta no existe en destino
        '...la copiamos con todo su contenido
        copyFolder path1 & folder.Name,path2 & folder.Name
      end if
    Next 'folder
    'En el diccionario sólo quedan referencias a carpetas que no existen en el origen
    'Debemos eliminarlas
    'Recorremos el diccionario
    for Each key in dic2.Keys
      'Eliminamos la carpeta de destino
      deleteFolder(path2 & key)
    Next
  End If
End Function

function pad2(a)
  'Ajusta un valor a 2 dígitos
   pad2=right("0" & a,2)
end function


'Programa principal
'Inicialmente no tenemos errores
errorCount=0
'Creamos el nombre del archivo de log en base a la fecha actual
n=now 'Anotamos la fecha y hora actuales
logFileName=DatePart("yyyy",n) & pad2(DatePart("m",n)) & pad2(DatePart("d",n))
logFileName=logFileName & pad2(DatePart("h",n)) & pad2(DatePart("n",n)) & pad2(DatePart("s",n))
logFileName=logFolder & logFileName & ".txt"
'Abrimos el archivo de log
Set logFile=fso.CreateTextFile(logFileName,true)
'Escribimos el inicio de log
logFile.write(n & " Sincronización de carpeta " & sourceFolder & " con " & destinationFolder & vbCrLf)
'Sincronizamos las carpetas
syncFolder sourceFolder,destinationFolder
'Finalizamos el log indicando el número de errores
logFile.write(now & " Sincronización finalizada con " & errorCount & " errores")
'Cerramos el archivo de log
logFile.Close