Attribute VB_Name = "MdiUsuarios"
Option Explicit

Public Function NombreModulo(ByVal Codigo As Integer) As String
    NombreModulo = Choose(Codigo, "Mantenimiento de Clientes", "Visualizar Informacion de Registros", "Generar Registros y Etiquetas", "Gestion de Servicios por Cliente", "Registro de Reparto y Ranches", "Generar Consolidados", "Confirmacion de Registros", "Control de Documentos a Delivery", "Informe de Status", "Configuracion de Informaciones", "Actualizaciones")
End Function

Public Function TieneAcceso(ByVal Modulo As Integer) As Boolean
    Dim Datos As ADODB.Recordset, Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Not EjecuteDatrav Then Exit Function
    Set Datos = ConsultarDatos("SELECT u.[Access],a.CodMod,a.Permitido FROM dbo.MdbUsers u LEFT JOIN dbo.IntAccesos a ON a.CodUsu=u.ITEM WHERE u.Usuari=? AND u.SitUsu=1 AND u.ActPas=1", VarUsuario, True)
    Do Until Datos.EOF
        If CLng(Datos.Fields(0).Value) = 0 Then TieneAcceso = True: Exit Do
        If Not IsNull(Datos.Fields(1).Value) Then
            If CLng(Datos.Fields(1).Value) = Modulo And CBool(Datos.Fields(2).Value) Then TieneAcceso = True: Exit Do
        End If
        Datos.MoveNext
    Loop
    Datos.Close
    Exit Function
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    On Error GoTo 0
    Err.Raise Numero, "TieneAcceso", Descripcion
End Function

Public Sub VerificarAccesoFormulario(ByVal Formulario As Form, ByVal Modulo As Integer)
    On Error GoTo ErrInt
    If TieneAcceso(Modulo) Then Exit Sub
    Formulario.Hide
    MsgBox "Su cuenta no tiene acceso a este modulo.", vbExclamation, "Clever SoftPro"
    Exit Sub
ErrInt:
    Dim Descripcion As String
    Descripcion = Err.Description
    Formulario.Hide
    MsgBox Descripcion, vbExclamation, "Clever SoftPro"
End Sub
Public Sub ExigirAdministrador()
    If Not TieneAcceso(0) Then Err.Raise 5, , "Solo un administrador activo puede gestionar usuarios."
End Sub

Public Function ListarUsuarios() As ADODB.Recordset
    ExigirAdministrador
    Set ListarUsuarios = ConsultarDatos("SELECT ITEM,NomUsu,ApeUsu,Usuari,[Access],SitUsu,ActPas,Email FROM dbo.MdbUsers ORDER BY Usuari")
End Function

Public Function GuardarUsuario(ByVal Codigo As Long, ByVal Accion As Integer, Optional ByVal Nombre As String = "", Optional ByVal Apellido As String = "", Optional ByVal Usuario As String = "", Optional ByVal Clave As String = "", Optional ByVal Confirmacion As String = "", Optional ByVal Perfil As Integer = 1, Optional ByVal Activo As Boolean = True, Optional ByVal Correo As String = "", Optional ByVal Permisos As String = "") As Long
    Dim Conexion As ADODB.Connection, Comando As ADODB.Command, Datos As ADODB.Recordset, Digest As MD5
    Dim Numero As Long, Descripcion As String, Resumen As String, Posicion As Integer
    On Error GoTo ErrInt
    ExigirAdministrador
    Nombre = Trim$(Nombre): Apellido = Trim$(Apellido): Usuario = UCase$(Trim$(Usuario)): Correo = Trim$(Correo)
    If Accion = 1 Or Accion = 2 Then
        If Len(Nombre) = 0 Or Len(Nombre) > 50 Or Len(Apellido) = 0 Or Len(Apellido) > 50 Then Err.Raise 5, , "Nombres y apellidos son obligatorios y admiten hasta 50 caracteres."
        If Perfil < 0 Or Perfil > 1 Or Len(Correo) > 150 Then Err.Raise 5, , "Revise el perfil y el correo (hasta 150 caracteres)."
    End If
    If Accion = 1 Then
        If Len(Usuario) < 1 Or Len(Usuario) > 5 Then Err.Raise 5, , "El usuario admite de 1 a 5 letras o numeros."
        For Posicion = 1 To Len(Usuario)
            If Not Mid$(Usuario, Posicion, 1) Like "[A-Z0-9]" Then Err.Raise 5, , "El usuario admite solo letras y numeros."
        Next
    End If
    If Accion = 1 Or Accion = 4 Then
        If Len(Clave) < 6 Or Len(Clave) > 150 Or Len(Trim$(Clave)) = 0 Then Err.Raise 5, , "La clave admite de 6 a 150 caracteres."
        If Clave <> Confirmacion Then Err.Raise 5, , "La clave y su confirmacion son distintas."
        Set Digest = New MD5: Resumen = LCase$(Digest.DigestStrToHexStr(Clave))
    End If
    Set Conexion = AbrirConexion(): Set Comando = New ADODB.Command
    Set Comando.ActiveConnection = Conexion
    Comando.CommandType = adCmdStoredProc: Comando.CommandText = "dbo.GuardarUsuario"
    Comando.Parameters.Append Comando.CreateParameter("@Actor", adInteger, adParamInput, , UsuarioOperativo())
    Comando.Parameters.Append Comando.CreateParameter("@Codigo", adInteger, adParamInput, , Codigo)
    Comando.Parameters.Append Comando.CreateParameter("@Accion", adInteger, adParamInput, , Accion)
    Comando.Parameters.Append Comando.CreateParameter("@Nombre", adVarWChar, adParamInput, 50, Nombre)
    Comando.Parameters.Append Comando.CreateParameter("@Apellido", adVarWChar, adParamInput, 50, Apellido)
    Comando.Parameters.Append Comando.CreateParameter("@Usuario", adVarWChar, adParamInput, 5, Usuario)
    Comando.Parameters.Append Comando.CreateParameter("@Clave", adVarChar, adParamInput, 32, Resumen)
    Comando.Parameters.Append Comando.CreateParameter("@Perfil", adUnsignedTinyInt, adParamInput, , Perfil)
    Comando.Parameters.Append Comando.CreateParameter("@Activo", adBoolean, adParamInput, , Activo)
    Comando.Parameters.Append Comando.CreateParameter("@Correo", adVarWChar, adParamInput, 150, Correo)
    If Accion = 5 Then Comando.Parameters.Append Comando.CreateParameter("@Permisos", adLongVarWChar, adParamInput, Len(Permisos), Permisos)
    Set Datos = Comando.Execute
    GuardarUsuario = CLng(Datos.Fields(0).Value)
    Datos.Close: Conexion.Close
    Exit Function
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    If Not Conexion Is Nothing Then Conexion.Close
    On Error GoTo 0
    Err.Raise Numero, "GuardarUsuario", Descripcion
End Function
