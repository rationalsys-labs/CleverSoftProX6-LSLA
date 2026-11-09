Attribute VB_Name = "MdiSesion"
Option Explicit

Public Function CargarUsuarios() As ADODB.Recordset
    Dim cn As ADODB.Connection
    Dim Datos As ADODB.Recordset
    Dim Numero As Long
    Dim Descripcion As String
    On Error GoTo ErrInt
    Set cn = AbrirConexion()
    Set Datos = New ADODB.Recordset
    Datos.CursorLocation = adUseClient
    Datos.Open "SELECT Usuari FROM dbo.MdbUsers WHERE SitUsu=1 AND ActPas=1 ORDER BY Usuari", cn, adOpenStatic, adLockReadOnly
    Set Datos.ActiveConnection = Nothing
    cn.Close
    Set CargarUsuarios = Datos
    Exit Function
ErrInt:
    Numero = Err.Number
    Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then
        If Datos.State <> adStateClosed Then Datos.Close
    End If
    If Not cn Is Nothing Then
        If cn.State <> adStateClosed Then cn.Close
    End If
    On Error GoTo 0
    Err.Raise Numero, "CargarUsuarios", Descripcion
End Function

Public Function AutenticarUsuario(ByVal Usuario As String, ByVal Pasword As String) As Boolean
    Dim cn As ADODB.Connection
    Dim Comando As ADODB.Command
    Dim Datos As ADODB.Recordset
    Dim Digest As MD5
    Dim Numero As Long
    Dim Descripcion As String
    On Error GoTo ErrInt
    'El procedimiento original acepta usuarios de cinco caracteres.
    If Len(Usuario) = 0 Or Len(Usuario) > 5 Then Exit Function
    Set Digest = New MD5
    Set cn = AbrirConexion()
    Set Comando = New ADODB.Command
    Set Comando.ActiveConnection = cn
    Comando.CommandType = adCmdStoredProc
    Comando.CommandText = "dbo.Valid_LoginClev"
    Comando.Parameters.Append Comando.CreateParameter("@Usuario", adVarChar, adParamInput, 5, Usuario)
    Comando.Parameters.Append Comando.CreateParameter("@Pasword", adVarChar, adParamInput, 150, LCase$(Digest.DigestStrToHexStr(Pasword)))
    Set Datos = Comando.Execute
    If Not Datos Is Nothing Then
        If Datos.State = adStateOpen Then
            If Not Datos.EOF Then
                If Not IsNull(Datos.Fields(0).Value) And Not IsNull(Datos.Fields(1).Value) Then
                    AutenticarUsuario = (CStr(Datos.Fields(0).Value) = "MenSQL" And CStr(Datos.Fields(1).Value) = "1")
                End If
            End If
            Datos.Close
        End If
    End If
    cn.Close
    Exit Function
ErrInt:
    Numero = Err.Number
    Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then
        If Datos.State <> adStateClosed Then Datos.Close
    End If
    If Not cn Is Nothing Then
        If cn.State <> adStateClosed Then cn.Close
    End If
    On Error GoTo 0
    Err.Raise Numero, "AutenticarUsuario", Descripcion
End Function

Public Function UsuarioOperativo() As Long
    Dim Datos As ADODB.Recordset
    Set Datos = ConsultarDatos("SELECT ITEM FROM dbo.MdbUsers WHERE Usuari=? AND SitUsu=1 AND ActPas=1", VarUsuario, True)
    If Datos.EOF Then Datos.Close: Err.Raise 5, , "El usuario de la sesion no esta disponible."
    UsuarioOperativo = CLng(Datos.Fields(0).Value)
    Datos.Close
End Function
