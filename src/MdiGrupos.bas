Attribute VB_Name = "MdiGrupos"
Option Explicit

Public Function CrearGrupo(ByVal Nombre As String) As Long
    Dim Conexion As ADODB.Connection, Comando As ADODB.Command, Datos As ADODB.Recordset
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    Nombre = UCase$(Trim$(Nombre))
    If Len(Nombre) = 0 Or Len(Nombre) > 150 Then Err.Raise 5, , "El nombre debe tener entre 1 y 150 caracteres."
    Set Conexion = AbrirConexion()
    Set Comando = New ADODB.Command
    Set Comando.ActiveConnection = Conexion
    Comando.CommandType = adCmdStoredProc: Comando.CommandText = "dbo.CrearGrupo"
    Comando.Parameters.Append Comando.CreateParameter("@Nombre", adVarWChar, adParamInput, 150, Nombre)
    Set Datos = Comando.Execute
    CrearGrupo = CLng(Datos.Fields(0).Value)
    Datos.Close: Conexion.Close
    Exit Function
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    If Not Conexion Is Nothing Then Conexion.Close
    On Error GoTo 0
    Err.Raise Numero, "CrearGrupo", Descripcion
End Function

Public Sub AgruparServicios(ByVal Grupo As Long, ByVal Accion As Integer, ByVal Codigos As Collection)
    Dim Conexion As ADODB.Connection, Comando As ADODB.Command
    Dim Codigo As Variant, Lista As String
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Grupo <= 0 Or Codigos.Count = 0 Then Err.Raise 5, , "Seleccione un grupo y sus servicios."
    Lista = "<Servicios>"
    For Each Codigo In Codigos
        If Not CodigoValido(CStr(Codigo)) Then Err.Raise 5, , "Codigo de servicio no valido."
        Lista = Lista & "<Servicio Codigo=""" & CStr(CDec(Codigo)) & """/>"
    Next
    Lista = Lista & "</Servicios>"
    Set Conexion = AbrirConexion()
    Set Comando = New ADODB.Command
    Set Comando.ActiveConnection = Conexion
    Comando.CommandType = adCmdStoredProc: Comando.CommandText = "dbo.AgruparServicios"
    Comando.Parameters.Append Comando.CreateParameter("@Grupo", adInteger, adParamInput, , Grupo)
    Comando.Parameters.Append Comando.CreateParameter("@Accion", adUnsignedTinyInt, adParamInput, , Accion)
    Comando.Parameters.Append Comando.CreateParameter("@Servicios", adLongVarWChar, adParamInput, Len(Lista), Lista)
    Comando.Execute , , adExecuteNoRecords
    Conexion.Close
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Conexion Is Nothing Then Conexion.Close
    On Error GoTo 0
    Err.Raise Numero, "AgruparServicios", Descripcion
End Sub
