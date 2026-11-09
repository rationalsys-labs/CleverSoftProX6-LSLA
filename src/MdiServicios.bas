Attribute VB_Name = "MdiServicios"
Option Explicit

Public Type DatosDocumento
    Servicio As String
    Cliente As String
    Producto As Long
    Referencia As String
    Destinatario As String
    Direccion As String
    Distrito As Long
    Contacto As String
    Descripcion As String
    CodigoDestinatario As String
End Type

Private Function ComandoServicio(ByVal Conexion As ADODB.Connection, ByVal Nombre As String) As ADODB.Command
    Dim Comando As ADODB.Command
    Set Comando = New ADODB.Command
    Set Comando.ActiveConnection = Conexion
    Comando.CommandType = adCmdStoredProc
    Comando.CommandText = "dbo." & Nombre
    Set ComandoServicio = Comando
End Function

Private Sub TextoParametro(ByVal Comando As ADODB.Command, ByVal Nombre As String, ByVal Longitud As Long, ByVal Valor As String)
    If Len(Valor) > Longitud Then Err.Raise vbObjectError + 301, , "El campo " & Nombre & " excede la longitud permitida."
    Comando.Parameters.Append Comando.CreateParameter(Nombre, adVarWChar, adParamInput, Longitud, Valor)
End Sub

Public Function CrearServicio(ByVal Cliente As String, ByVal Tipo As Long, ByVal Cantidad As String, ByVal Producto As Long, ByVal Boleta As String) As String
    Dim Conexion As ADODB.Connection
    Dim Comando As ADODB.Command
    Dim Datos As ADODB.Recordset
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Not CodigoValido(Cantidad) Then Err.Raise vbObjectError + 302, , "La cantidad debe ser un entero positivo."
    If CDec(Cantidad) > 2147483647# Then Err.Raise vbObjectError + 303, , "La cantidad supera el limite admitido por el formulario."
    Set Conexion = AbrirConexion()
    Set Comando = ComandoServicio(Conexion, "SaveNewService")
    AgregarCodigo Comando, "@Cliente", Cliente
    Comando.Parameters.Append Comando.CreateParameter("@TipoServicio", adInteger, adParamInput, , Tipo)
    AgregarCodigo Comando, "@Cantidad", Cantidad
    Comando.Parameters.Append Comando.CreateParameter("@Producto", adInteger, adParamInput, , Producto)
    AgregarCodigo Comando, "@Boleta", Boleta
    Set Datos = Comando.Execute
    CrearServicio = CStr(Datos.Fields(0).Value)
    Datos.Close: Conexion.Close
    Exit Function
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    If Not Conexion Is Nothing Then Conexion.Close
    On Error GoTo 0
    Err.Raise Numero, "CrearServicio", Descripcion
End Function

Public Function GuardarRegistro(ByRef Documento As DatosDocumento, Optional ByVal ConexionCarga As ADODB.Connection = Nothing) As String
    Dim Conexion As ADODB.Connection
    Dim Comando As ADODB.Command
    Dim Datos As ADODB.Recordset
    Dim Propia As Boolean
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    Propia = (ConexionCarga Is Nothing)
    If Propia Then Set Conexion = AbrirConexion() Else Set Conexion = ConexionCarga
    Set Comando = ComandoServicio(Conexion, "SaveNewRegister")
    AgregarCodigo Comando, "@Servicio", Documento.Servicio
    TextoParametro Comando, "@Documento", 50, Documento.Referencia
    Comando.Parameters.Append Comando.CreateParameter("@Producto", adInteger, adParamInput, , Documento.Producto)
    AgregarCodigo Comando, "@Cliente", Documento.Cliente
    TextoParametro Comando, "@Destinatario", 500, Documento.Destinatario
    TextoParametro Comando, "@Direccion", 500, Documento.Direccion
    Comando.Parameters.Append Comando.CreateParameter("@Distrito", adInteger, adParamInput, , Documento.Distrito)
    TextoParametro Comando, "@Contacto", 150, Documento.Contacto
    TextoParametro Comando, "@Descrip", 150, Documento.Descripcion
    If Documento.CodigoDestinatario = "" Then Documento.CodigoDestinatario = "0"
    AgregarCodigo Comando, "@CodSave", Documento.CodigoDestinatario
    Set Datos = Comando.Execute
    GuardarRegistro = CStr(Datos.Fields(0).Value)
    Datos.Close
    If Propia Then Conexion.Close
    Exit Function
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    If Propia And Not Conexion Is Nothing Then Conexion.Close
    On Error GoTo 0
    Err.Raise Numero, "GuardarRegistro", Descripcion
End Function

Public Sub CambiarServicio(ByVal Codigo As String, ByVal Accion As Integer, Optional ByVal Cantidad As String = "0")
    Dim Conexion As ADODB.Connection
    Dim Comando As ADODB.Command
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Accion = 1 Then
        If Not CodigoValido(Cantidad) Then Err.Raise vbObjectError + 302, , "La cantidad debe ser un entero positivo."
        If CDec(Cantidad) > 2147483647# Then Err.Raise vbObjectError + 303, , "La cantidad supera el limite admitido por el formulario."
    End If
    Set Conexion = AbrirConexion()
    Set Comando = ComandoServicio(Conexion, "ActualizarServicio")
    AgregarCodigo Comando, "@Codigo", Codigo
    Comando.Parameters.Append Comando.CreateParameter("@Accion", adUnsignedTinyInt, adParamInput, , Accion)
    AgregarCodigo Comando, "@Cantidad", Cantidad
    Comando.Execute , , adExecuteNoRecords
    Conexion.Close
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Conexion Is Nothing Then Conexion.Close
    On Error GoTo 0
    Err.Raise Numero, "CambiarServicio", Descripcion
End Sub

Public Sub AnularRegistro(ByVal Codigo As String)
    Dim Conexion As ADODB.Connection
    Dim Comando As ADODB.Command
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    Set Conexion = AbrirConexion()
    Set Comando = ComandoServicio(Conexion, "AnularDocumento")
    AgregarCodigo Comando, "@Codigo", Codigo
    Comando.Execute , , adExecuteNoRecords
    Conexion.Close
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Conexion Is Nothing Then Conexion.Close
    On Error GoTo 0
    Err.Raise Numero, "AnularRegistro", Descripcion
End Sub

Public Function CodigoValido(ByVal Texto As String) As Boolean
    Dim Posicion As Long
    If Len(Texto) = 0 Or Len(Texto) > 18 Then Exit Function
    For Posicion = 1 To Len(Texto)
        If Mid$(Texto, Posicion, 1) < "0" Or Mid$(Texto, Posicion, 1) > "9" Then Exit Function
    Next
    CodigoValido = (CDec(Texto) > 0)
End Function

Public Sub CargarLista(ByVal Lista As ComboBox, ByVal Consulta As String)
    Dim Datos As ADODB.Recordset
    Set Datos = ConsultarDatos(Consulta)
    Lista.Clear
    Do Until Datos.EOF
        Lista.AddItem CStr(Datos.Fields(1).Value)
        Lista.ItemData(Lista.NewIndex) = CLng(Datos.Fields(0).Value)
        Datos.MoveNext
    Loop
    Datos.Close
End Sub
