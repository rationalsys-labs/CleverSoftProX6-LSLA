Attribute VB_Name = "MdiActualizaciones"
Option Explicit
Private Const ENCABEZADO As String = "-- Clever SoftPro X6 - LSLA - Esquema 2"

Private Sub SaltarEspacios(ByVal Texto As String, ByRef Posicion As Long)
    Do While Posicion <= Len(Texto)
        If InStr(1, " " & vbTab & vbCr & vbLf, Mid$(Texto, Posicion, 1), vbBinaryCompare) = 0 Then Exit Do
        Posicion = Posicion + 1
    Loop
End Sub

Private Sub Exigir(ByVal Texto As String, ByRef Posicion As Long, ByVal Esperado As String)
    SaltarEspacios Texto, Posicion
    If StrComp(Mid$(Texto, Posicion, Len(Esperado)), Esperado, vbTextCompare) <> 0 Then Err.Raise 5, , "El archivo contiene una sentencia de actualizacion no admitida."
    Posicion = Posicion + Len(Esperado)
End Sub

Private Function LeerLiteral(ByVal Texto As String, ByRef Posicion As Long) As String
    Dim Inicio As Long, Final As Long, Resultado As String
    Exigir Texto, Posicion, "N'"
    Inicio = Posicion
    Do
        Final = InStr(Posicion, Texto, "'", vbBinaryCompare)
        If Final = 0 Then Err.Raise 5, , "El archivo contiene un texto SQL incompleto."
        Resultado = Resultado & Mid$(Texto, Inicio, Final - Inicio)
        Posicion = Final + 1
        If Mid$(Texto, Posicion, 1) <> "'" Then Exit Do
        Resultado = Resultado & "'": Posicion = Posicion + 1: Inicio = Posicion
    Loop
    LeerLiteral = Resultado
End Function

Public Function AnalizarActualizaciones(ByVal Texto As String) As Collection
    Dim Resultado As New Collection, Posicion As Long, Tabla As String, Datos As String
    If Left$(Texto, Len(ENCABEZADO)) <> ENCABEZADO Then Err.Raise 5, , "Use un archivo de actualizacion del esquema 2 de Clever."
    Posicion = Len(ENCABEZADO) + 1
    If Mid$(Texto, Posicion, 2) <> vbCrLf And Mid$(Texto, Posicion, 1) <> vbLf Then Err.Raise 5, , "El encabezado de actualizacion esta incompleto."
    Do
        SaltarEspacios Texto, Posicion
        If Posicion > Len(Texto) Then Exit Do
        Exigir Texto, Posicion, "EXEC": Exigir Texto, Posicion, "dbo.AplicarActualizacion"
        Exigir Texto, Posicion, "@Tabla": Exigir Texto, Posicion, "="
        Tabla = LeerLiteral(Texto, Posicion)
        Select Case Tabla
            Case "MDBCostumer", "MDBServices", "MDBDocument", "MDBHistory", "MDBHisEmp", "IntCorrel", "MdbGruposServ", "IntStatus"
            Case Else: Err.Raise 5, , "El archivo contiene una tabla de actualizacion no admitida."
        End Select
        Exigir Texto, Posicion, ",": Exigir Texto, Posicion, "@Datos": Exigir Texto, Posicion, "="
        Datos = LeerLiteral(Texto, Posicion)
        Exigir Texto, Posicion, ";"
        Resultado.Add Array(Tabla, Datos)
    Loop
    If Resultado.Count = 0 Then Err.Raise 5, , "El archivo no contiene actualizaciones."
    Set AnalizarActualizaciones = Resultado
End Function

Public Function ExportarActualizaciones(ByVal Ruta As String, Optional ByVal Sobrescribir As Boolean = False) As Long
    Dim Datos As ADODB.Recordset, Conexion As ADODB.Connection, Comando As ADODB.Command
    Dim Archivos As Object, Archivo As Object, Temporal As String, Seleccion As String, Publicado As Boolean
    Dim Numero As Long, Descripcion As String, Cantidad As Long, Validados As Collection
    On Error GoTo ErrInt
    Set Datos = ConsultarDatos("SELECT Item,Consulta FROM dbo.IntSaveActualizar WHERE VersionEsquema=2 AND ActSav=1 ORDER BY Item")
    If Datos.EOF Then Err.Raise 5, , "No existen actualizaciones pendientes del esquema actual."
    Temporal = TemporalDestino(Ruta, Sobrescribir)
    Set Archivos = CreateObject("Scripting.FileSystemObject")
    Set Archivo = Archivos.CreateTextFile(Temporal, False, True)
    Archivo.WriteLine ENCABEZADO
    Seleccion = "<Actualizaciones>"
    Do Until Datos.EOF
        Set Validados = AnalizarActualizaciones(ENCABEZADO & vbCrLf & CStr(Datos.Fields("Consulta").Value))
        If Validados.Count <> 1 Then Err.Raise 5, , "Una anotacion contiene varias actualizaciones."
        Archivo.WriteLine CStr(Datos.Fields("Consulta").Value)
        Seleccion = Seleccion & "<Item Codigo=""" & CStr(Datos.Fields("Item").Value) & """/>"
        Cantidad = Cantidad + 1: Datos.MoveNext
    Loop
    Seleccion = Seleccion & "</Actualizaciones>"
    Datos.Close: Set Datos = Nothing: Archivo.Close: Set Archivo = Nothing
    PublicarArchivo Temporal, Ruta, Sobrescribir
    Temporal = "": Publicado = True
    Set Conexion = AbrirConexion(): Set Comando = New ADODB.Command
    Set Comando.ActiveConnection = Conexion
    Comando.CommandType = adCmdStoredProc: Comando.CommandText = "dbo.MarcarActualizaciones"
    Comando.Parameters.Append Comando.CreateParameter("@Seleccion", adLongVarWChar, adParamInput, Len(Seleccion), Seleccion)
    Comando.Execute , , adExecuteNoRecords
    Conexion.Close
    ExportarActualizaciones = Cantidad
    Exit Function
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Archivo Is Nothing Then Archivo.Close
    If Not Datos Is Nothing Then Datos.Close
    If Not Conexion Is Nothing Then Conexion.Close
    QuitarTemporal Temporal
    Err.Clear
    On Error GoTo 0
    If Publicado Then Descripcion = "El archivo fue generado, pero no se pudo confirmar el marcado de la cola. " & Descripcion
    Err.Raise Numero, "ExportarActualizaciones", Descripcion
End Function

Public Function ImportarActualizaciones(ByVal Ruta As String) As Long
    Dim Archivos As Object, Archivo As Object, Lote As Collection, Entrada As Variant
    Dim Conexion As ADODB.Connection, Comando As ADODB.Command, EnTransaccion As Boolean
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    Set Archivos = CreateObject("Scripting.FileSystemObject")
    Set Archivo = Archivos.OpenTextFile(Ruta, 1, False, -1)
    Set Lote = AnalizarActualizaciones(Archivo.ReadAll)
    Archivo.Close: Set Archivo = Nothing
    Set Conexion = AbrirConexion()
    Conexion.BeginTrans: EnTransaccion = True
    For Each Entrada In Lote
        Set Comando = New ADODB.Command
        Set Comando.ActiveConnection = Conexion
        Comando.CommandType = adCmdStoredProc: Comando.CommandText = "dbo.AplicarActualizacion"
        Comando.Parameters.Append Comando.CreateParameter("@Tabla", adVarWChar, adParamInput, 128, CStr(Entrada(0)))
        Comando.Parameters.Append Comando.CreateParameter("@Datos", adLongVarWChar, adParamInput, Len(CStr(Entrada(1))), CStr(Entrada(1)))
        Comando.Execute , , adExecuteNoRecords
    Next
    Conexion.CommitTrans: EnTransaccion = False
    Conexion.Close
    ImportarActualizaciones = Lote.Count
    Exit Function
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Archivo Is Nothing Then Archivo.Close
    If Not Conexion Is Nothing Then
        If EnTransaccion Then Conexion.RollbackTrans
        Conexion.Close
    End If
    Err.Clear
    On Error GoTo 0
    Err.Raise Numero, "ImportarActualizaciones", Descripcion
End Function
