Attribute VB_Name = "MdiConfiguracion"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Function LeerConexion(ByVal Clave As String, ByVal Predeterminado As String) As String
    Dim Buffer As String
    Dim Longitud As Long
    Dim Ruta As String
    Ruta = App.Path
    If Right$(Ruta, 1) <> "\" Then Ruta = Ruta & "\"
    Ruta = Ruta & "Config\Clever.local.ini"
    Buffer = Space$(2048)
    Longitud = GetPrivateProfileString("Conexion", Clave, Predeterminado, Buffer, Len(Buffer), Ruta)
    If Longitud = Len(Buffer) - 1 Then Err.Raise vbObjectError + 100, "LeerConexion", "Valor de configuracion demasiado largo."
    LeerConexion = Left$(Buffer, Longitud)
End Function

Public Function AbrirConexion() As ADODB.Connection
    Dim cn As ADODB.Connection
    Dim Servidor As String
    Dim Base As String
    Dim Modo As String
    Servidor = LeerConexion("Servidor", "localhost\SQLEXPRESS")
    Base = LeerConexion("Base", "DCSP")
    Modo = LeerConexion("AutenticacionWindows", "1")
    If Len(Trim$(Servidor)) = 0 Or Len(Trim$(Base)) = 0 Then Err.Raise vbObjectError + 101, "AbrirConexion", "Falta el servidor o la base de datos."
    If Modo <> "0" And Modo <> "1" Then Err.Raise vbObjectError + 102, "AbrirConexion", "AutenticacionWindows debe ser 0 o 1."
    Set cn = New ADODB.Connection
    cn.Provider = "SQLOLEDB.1"
    cn.ConnectionTimeout = 10
    cn.CommandTimeout = 30
    cn.CursorLocation = adUseClient
    cn.Properties("Data Source").Value = Servidor
    cn.Properties("Initial Catalog").Value = Base
    If Modo = "1" Then
        cn.Properties("Integrated Security").Value = "SSPI"
    Else
        cn.Properties("User ID").Value = LeerConexion("Usuario", vbNullString)
        cn.Properties("Password").Value = LeerConexion("Clave", vbNullString)
    End If
    cn.Open
    Set AbrirConexion = cn
End Function
