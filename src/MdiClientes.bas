Attribute VB_Name = "MdiClientes"
Option Explicit

Public Type DatosCliente
    Codigo As String
    Nombre As String
    TipoDocumento As Long
    Documento As String
    Direccion As String
    Distrito As Long
    Representante As String
    CorreoRepresentante As String
    Contacto As String
    Telefono As String
    TelefonoTrabajo As String
    TelefonoMovil As String
    TelefonoFax As String
    CorreoContacto As String
    Activo As Boolean
End Type

Public Function ConsultarDatos(ByVal Consulta As String, Optional ByVal Busqueda As String = "", Optional ByVal Filtrar As Boolean = False) As ADODB.Recordset
    Dim Conexion As ADODB.Connection
    Dim Comando As ADODB.Command
    Dim Datos As ADODB.Recordset
    Dim Numero As Long
    Dim Descripcion As String
    On Error GoTo ErrInt
    Set Conexion = AbrirConexion()
    Set Comando = New ADODB.Command
    Set Comando.ActiveConnection = Conexion
    Comando.CommandType = adCmdText
    Comando.CommandText = Consulta
    If Filtrar Then Comando.Parameters.Append Comando.CreateParameter("Busqueda", adVarWChar, adParamInput, 1000, Busqueda)
    Set Datos = New ADODB.Recordset
    Datos.CursorLocation = adUseClient
    Datos.Open Comando, , adOpenStatic, adLockReadOnly
    Set Datos.ActiveConnection = Nothing
    Conexion.Close
    Set ConsultarDatos = Datos
    Exit Function
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then
        If Datos.State <> adStateClosed Then Datos.Close
    End If
    If Not Conexion Is Nothing Then
        If Conexion.State <> adStateClosed Then Conexion.Close
    End If
    On Error GoTo 0
    Err.Raise Numero, "ConsultarDatos", Descripcion
End Function

Public Function TextoDato(ByVal Valor As Variant) As String
    If Not IsNull(Valor) Then TextoDato = CStr(Valor)
End Function

Public Sub AgregarCodigo(ByVal Comando As ADODB.Command, ByVal Nombre As String, ByVal Codigo As String)
    Dim Parametro As ADODB.Parameter
    Set Parametro = Comando.CreateParameter(Nombre, adNumeric, adParamInput)
    Parametro.Precision = 18
    Parametro.NumericScale = 0
    Parametro.Value = CDec(Codigo)
    Comando.Parameters.Append Parametro
End Sub

Private Sub AgregarTexto(ByVal Comando As ADODB.Command, ByVal Nombre As String, ByVal Longitud As Long, ByVal Valor As String)
    If Len(Valor) > Longitud Then Err.Raise vbObjectError + 201, , "El campo " & Nombre & " excede su longitud."
    Comando.Parameters.Append Comando.CreateParameter(Nombre, adVarWChar, adParamInput, Longitud, Valor)
End Sub

Public Function GuardarCliente(ByRef Cliente As DatosCliente, ByVal Modalidad As Integer) As String
    Dim Conexion As ADODB.Connection
    Dim Comando As ADODB.Command
    Dim Datos As ADODB.Recordset
    Dim Numero As Long
    Dim Descripcion As String
    On Error GoTo ErrInt
    Set Conexion = AbrirConexion()
    Set Comando = New ADODB.Command
    Set Comando.ActiveConnection = Conexion
    Comando.CommandType = adCmdStoredProc
    Comando.CommandText = "dbo.SaveNewModClient"
    AgregarCodigo Comando, "@Codigo", Cliente.Codigo
    AgregarTexto Comando, "@Cliente", 300, Cliente.Nombre
    Comando.Parameters.Append Comando.CreateParameter("@Tipo", adInteger, adParamInput, , Cliente.TipoDocumento)
    AgregarTexto Comando, "@RucDni", 20, Cliente.Documento
    AgregarTexto Comando, "@Direccion", 300, Cliente.Direccion
    Comando.Parameters.Append Comando.CreateParameter("@Distrito", adInteger, adParamInput, , Cliente.Distrito)
    AgregarTexto Comando, "@Representante", 150, Cliente.Representante
    AgregarTexto Comando, "@MailRepre", 150, Cliente.CorreoRepresentante
    AgregarTexto Comando, "@Contacto", 150, Cliente.Contacto
    AgregarTexto Comando, "@Telefono", 50, Cliente.Telefono
    AgregarTexto Comando, "@Teltrabajo", 50, Cliente.TelefonoTrabajo
    AgregarTexto Comando, "@TelMovil", 50, Cliente.TelefonoMovil
    AgregarTexto Comando, "@TelFax", 50, Cliente.TelefonoFax
    AgregarTexto Comando, "@EmailCont", 150, Cliente.CorreoContacto
    Comando.Parameters.Append Comando.CreateParameter("@Activate", adBoolean, adParamInput, , Cliente.Activo)
    Comando.Parameters.Append Comando.CreateParameter("@Modal", adUnsignedTinyInt, adParamInput, , Modalidad)
    Set Datos = Comando.Execute
    If Datos.EOF Then Err.Raise vbObjectError + 202, , "No se recibio el codigo del cliente."
    GuardarCliente = CStr(Datos.Fields(0).Value)
    Datos.Close
    Conexion.Close
    Exit Function
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then
        If Datos.State <> adStateClosed Then Datos.Close
    End If
    If Not Conexion Is Nothing Then
        If Conexion.State <> adStateClosed Then Conexion.Close
    End If
    On Error GoTo 0
    Err.Raise Numero, "GuardarCliente", Descripcion
End Function

Public Sub EliminarCliente(ByVal Codigo As String)
    Dim Conexion As ADODB.Connection
    Dim Comando As ADODB.Command
    Dim Numero As Long
    Dim Descripcion As String
    Dim EnTransaccion As Boolean
    On Error GoTo ErrInt
    Set Conexion = AbrirConexion()
    Conexion.BeginTrans
    EnTransaccion = True
    Set Comando = New ADODB.Command
    Set Comando.ActiveConnection = Conexion
    Comando.CommandType = adCmdText
    Comando.CommandText = "SET NOCOUNT ON; DECLARE @Codigo numeric(18,0); SET @Codigo=?; " & _
        "UPDATE dbo.MDBCostumer SET Active=0 WHERE CodCli=@Codigo AND Active=1; " & _
        "IF @@ROWCOUNT<>1 RAISERROR('El cliente ya no esta disponible.',16,1); " & _
        "ELSE EXEC dbo.RegistrarActualizacion N'MDBCostumer',@Codigo;"
    AgregarCodigo Comando, "Codigo", Codigo
    Comando.Execute , , adExecuteNoRecords
    Conexion.CommitTrans
    EnTransaccion = False
    Conexion.Close
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Conexion Is Nothing Then
        If EnTransaccion Then Conexion.RollbackTrans
        If Conexion.State <> adStateClosed Then Conexion.Close
    End If
    On Error GoTo 0
    Err.Raise Numero, "EliminarCliente", Descripcion
End Sub

Public Function BuscarClientes(ByVal Texto As String) As ADODB.Recordset
    'Los comodines escritos por el usuario se interpretan como texto literal.
    Texto = Replace(Replace(Replace(Texto, "[", "[[]"), "%", "[%]"), "_", "[_]")
    Set BuscarClientes = ConsultarDatos("SELECT TOP 50 Asiento,Cliente,RUC,Direccion,Distrito FROM dbo.InforCostumer WHERE Cliente LIKE ? ORDER BY Cliente,Asiento", "%" & Texto & "%", True)
End Function

Public Function BuscarRegistros(ByVal Modo As Integer, ByVal Campo As String, ByVal Texto As String) As ADODB.Recordset
    Dim Consulta As String
    Dim Columna As String
    Select Case Modo
        Case 1
            Consulta = "SELECT Codigo,Boleta,Cliente,CodCliente,Producto,Cantidad FROM dbo.ViewServices"
            Select Case Campo
                Case "Codigo", "Boleta", "Cliente", "Producto": Columna = Campo
                Case Else: Columna = "Cliente"
            End Select
        Case 2
            Consulta = "SELECT TOP 50 Codigo,Empresa,Direccion2,Contacto,Servicio FROM dbo.ViewVisDocument"
            Select Case Campo
                Case "Codigo", "Empresa", "Direccion2", "Contacto", "Servicio": Columna = Campo
                Case Else: Columna = "Empresa"
            End Select
        Case 3
            Consulta = "SELECT TOP 51 Item,NomGrp FROM dbo.MdbGruposServ"
            If Campo = "Item" Then Columna = "Item" Else Columna = "NomGrp"
        Case Else
            Err.Raise vbObjectError + 204, , "Tipo de busqueda no valido."
    End Select
    If Len(Texto) > 0 Then
        If Columna = "Codigo" Or Columna = "Servicio" Or Columna = "Item" Or Columna = "Boleta" Then
            If IsNumeric(Texto) Then Texto = CStr(CDec(Texto))
        End If
        Texto = Replace(Replace(Replace(Texto, "[", "[[]"), "%", "[%]"), "_", "[_]")
        Consulta = Consulta & " WHERE CONVERT(nvarchar(500),[" & Columna & "]) LIKE ?"
    End If
    Consulta = Consulta & " ORDER BY 1 DESC"
    Set BuscarRegistros = ConsultarDatos(Consulta, "%" & Texto & "%", Len(Texto) > 0)
End Function
