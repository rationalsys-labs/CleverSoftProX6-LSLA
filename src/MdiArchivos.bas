Attribute VB_Name = "MdiArchivos"
Option Explicit
Private Declare Function RenombrarArchivo Lib "kernel32" Alias "MoveFileExW" (ByVal Origen As Long, ByVal Destino As Long, ByVal Opciones As Long) As Long

Public Function TemporalDestino(ByVal Ruta As String, Optional ByVal Sobrescribir As Boolean = False) As String
    Dim Archivos As Object, Carpeta As String
    Set Archivos = CreateObject("Scripting.FileSystemObject")
    If Len(Trim$(Ruta)) = 0 Then Err.Raise 5, , "Seleccione un archivo de destino."
    Carpeta = Archivos.GetParentFolderName(Archivos.GetAbsolutePathName(Ruta))
    If Not Archivos.FolderExists(Carpeta) Or Archivos.FolderExists(Ruta) Then Err.Raise 5, , "El destino no es un archivo valido."
    If Archivos.FileExists(Ruta) And Not Sobrescribir Then Err.Raise 5, , "El archivo ya existe. Seleccione otro destino."
    TemporalDestino = Archivos.BuildPath(Carpeta, Archivos.GetTempName)
End Function

Public Sub PublicarArchivo(ByVal Temporal As String, ByVal Ruta As String, Optional ByVal Sobrescribir As Boolean = False)
    Dim Opciones As Long
    Opciones = 8
    If Sobrescribir Then Opciones = Opciones Or 1
    If RenombrarArchivo(StrPtr(Temporal), StrPtr(Ruta), Opciones) = 0 Then Err.Raise vbObjectError + 600, , "No se pudo publicar el archivo de destino."
End Sub

Public Sub QuitarTemporal(ByVal Ruta As String)
    Dim Archivos As Object
    If Len(Ruta) = 0 Then Exit Sub
    Set Archivos = CreateObject("Scripting.FileSystemObject")
    If Archivos.FileExists(Ruta) Then Archivos.DeleteFile Ruta
End Sub
' Cada celda queda entre comillas para conservar tabs, saltos y comillas.
Public Function CeldaTexto(ByVal Texto As String) As String
    CeldaTexto = """" & Replace(Texto, """", """""") & """"
End Function
