Attribute VB_Name = "MdiReparto"
Option Explicit



Public Function RegistrarReparto(ByVal Cuadricula As Object, ByVal Mensajero As Long, ByVal Modalidad As Integer, ByRef Fecha As Date) As String
    Dim Conexion As ADODB.Connection, Comando As ADODB.Command, Datos As ADODB.Recordset
    Dim Lista As String, Fila As Long, Codigo As String
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Cuadricula.Rows < 2 Then Err.Raise 5, , "No existen documentos para repartir."
    Lista = "<Documentos>"
    For Fila = 1 To Cuadricula.Rows - 1
        Codigo = Cuadricula.TextMatrix(Fila, 0)
        If Not CodigoValido(Codigo) Then Err.Raise 5, , "Codigo de documento no valido."
        Lista = Lista & "<Documento Codigo=""" & CStr(CDec(Codigo)) & """/>"
    Next
    Lista = Lista & "</Documentos>"
    Set Conexion = AbrirConexion()
    Set Comando = New ADODB.Command
    Set Comando.ActiveConnection = Conexion
    Comando.CommandType = adCmdStoredProc: Comando.CommandText = "dbo.RegistrarReparto"
    Comando.Parameters.Append Comando.CreateParameter("@Usuario", adInteger, adParamInput, , UsuarioOperativo())
    Comando.Parameters.Append Comando.CreateParameter("@Mensajero", adInteger, adParamInput, , Mensajero)
    Comando.Parameters.Append Comando.CreateParameter("@Modalidad", adUnsignedTinyInt, adParamInput, , Modalidad)
    Comando.Parameters.Append Comando.CreateParameter("@Documentos", adLongVarWChar, adParamInput, Len(Lista), Lista)
    Set Datos = Comando.Execute
    RegistrarReparto = CStr(Datos.Fields("Ranche").Value): Fecha = CDate(Datos.Fields("Fecha").Value)
    Datos.Close: Conexion.Close
    Exit Function
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    If Not Conexion Is Nothing Then Conexion.Close
    On Error GoTo 0
    Err.Raise Numero, "RegistrarReparto", Descripcion
End Function

Public Sub LiberarGuia(ByVal Codigo As String)
    Dim Conexion As ADODB.Connection, Comando As ADODB.Command
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Not CodigoValido(Codigo) Then Err.Raise 5, , "Ingrese el codigo de la guia."
    Set Conexion = AbrirConexion()
    Set Comando = New ADODB.Command
    Set Comando.ActiveConnection = Conexion
    Comando.CommandType = adCmdStoredProc: Comando.CommandText = "dbo.LiberarGuia"
    AgregarCodigo Comando, "@Codigo", Codigo
    Comando.Parameters.Append Comando.CreateParameter("@Usuario", adInteger, adParamInput, , UsuarioOperativo())
    Comando.Execute , , adExecuteNoRecords
    Conexion.Close
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Conexion Is Nothing Then Conexion.Close
    On Error GoTo 0
    Err.Raise Numero, "LiberarGuia", Descripcion
End Sub

Public Sub AgregarDocumentoReparto(ByVal Cuadricula As Object, ByVal Codigo As String, ByVal Consolidado As Boolean, ByVal Mensajero As String)
    Dim Datos As ADODB.Recordset, Fila As Long, Columna As Integer
    If Not CodigoValido(Codigo) Then Err.Raise 5, , "Ingrese un codigo de documento valido."
    Codigo = CStr(CDec(Codigo))
    For Fila = 1 To Cuadricula.Rows - 1
        If Cuadricula.TextMatrix(Fila, 0) = Codigo Then Err.Raise 5, , "El documento ya se encuentra en la lista."
    Next
    Set Datos = ConsultarDatos("SELECT v.Codigo,v.Producto,v.Empresa,v.Direccion,v.Distrito,v.Contacto,v.Fecha,v.Cliente,v.Descripcion,v.Status,CASE WHEN EXISTS(SELECT 1 FROM dbo.MDBHistory h WHERE h.NroCod=d.NroCod AND h.Status=1 AND h.Active=1) THEN 1 ELSE 0 END Asignado FROM dbo.ViewDocument v JOIN dbo.MDBDocument d ON d.NroCod=v.Codigo WHERE d.Active=1 AND v.Codigo=?", Codigo, True)
    If Datos.EOF Then Datos.Close: Err.Raise 5, , "No se encontro el documento o fue anulado."
    If CLng(Datos.Fields("Asignado").Value) = 1 Then
        Datos.Close
        Err.Raise vbObjectError + 401, , "La guia ya esta asignada. Debe liberarla antes de volver a repartirla."
    End If
    Cuadricula.Rows = Cuadricula.Rows + 1: Fila = Cuadricula.Rows - 1
    For Columna = 0 To 6: Cuadricula.TextMatrix(Fila, Columna) = TextoDato(Datos.Fields(Columna).Value): Next
    If Consolidado Then
        Cuadricula.TextMatrix(Fila, 7) = TextoDato(Datos.Fields("Descripcion").Value)
        Cuadricula.TextMatrix(Fila, 8) = Mensajero
        Cuadricula.TextMatrix(Fila, 9) = TextoDato(Datos.Fields("Status").Value)
    Else
        Cuadricula.TextMatrix(Fila, 7) = Mensajero
        Cuadricula.TextMatrix(Fila, 8) = TextoDato(Datos.Fields("Cliente").Value)
    End If
    Datos.Close
End Sub

Public Sub PrepararCuadriculaReparto(ByVal Cuadricula As Object, ByVal Consolidado As Boolean)
    Dim Campos As Variant, Anchos As Variant, Columna As Integer
    Campos = Array("Codigo", "Producto", "Empresa", "Direccion", "Distrito", "Contacto", "Fecha", "Mensajero", "cliente")
    Anchos = Array(900, 1000, 3000, 2400, 1300, 1500, 0, 0, 0)
    If Consolidado Then
        Campos = Array("Codigo", "Producto", "Empresa", "Direccion", "Distrito", "Contacto", "Fecha", "Descripcion", "Mensajero", "Status")
        Anchos = Array(900, 1000, 3000, 2400, 1300, 1500, 0, 0, 1500, 1200)
    End If
    Cuadricula.Rows = 1: Cuadricula.Cols = UBound(Campos) + 1
    For Columna = 0 To UBound(Campos)
        Cuadricula.TextMatrix(0, Columna) = Campos(Columna): Cuadricula.ColWidth(Columna) = Anchos(Columna)
    Next
End Sub

Public Sub QuitarDocumentoReparto(ByVal Cuadricula As Object)
    If Cuadricula.Rows < 2 Or Cuadricula.Row < 1 Then Exit Sub
    If Cuadricula.Rows = 2 Then Cuadricula.Rows = 1 Else Cuadricula.RemoveItem Cuadricula.Row
End Sub