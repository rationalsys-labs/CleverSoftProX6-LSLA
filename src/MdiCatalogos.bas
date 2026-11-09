Attribute VB_Name = "MdiCatalogos"
Option Explicit

Public Function GuardarEstado(ByVal Codigo As Long, ByVal Accion As Integer, Optional ByVal Nombre As String = "") As Long
    Dim Conexion As ADODB.Connection, Comando As ADODB.Command, Datos As ADODB.Recordset
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Accion = 1 Or Accion = 2 Then
        Nombre = Trim$(Nombre)
        If Len(Nombre) = 0 Or Len(Nombre) > 20 Then Err.Raise 5, , "El estado admite entre 1 y 20 caracteres."
    End If
    Set Conexion = AbrirConexion(): Set Comando = New ADODB.Command
    Set Comando.ActiveConnection = Conexion
    Comando.CommandType = adCmdStoredProc: Comando.CommandText = "dbo.GuardarEstado"
    Comando.Parameters.Append Comando.CreateParameter("@Codigo", adInteger, adParamInput, , Codigo)
    Comando.Parameters.Append Comando.CreateParameter("@Accion", adUnsignedTinyInt, adParamInput, , Accion)
    Comando.Parameters.Append Comando.CreateParameter("@Nombre", adVarWChar, adParamInput, 20, Nombre)
    Set Datos = Comando.Execute
    GuardarEstado = CLng(Datos.Fields(0).Value)
    Datos.Close: Conexion.Close
    Exit Function
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    If Not Conexion Is Nothing Then Conexion.Close
    Err.Clear
    On Error GoTo 0
    Err.Raise Numero, "GuardarEstado", Descripcion
End Function
