Attribute VB_Name = "MdiExcel"
Option Explicit

Public Function ImportarDocumentos(ByVal Ruta As String, ByRef Plantilla As DatosDocumento) As Collection
    Dim Excel As Object, Libro As Object, Hoja As Object
    Dim Distritos As Object
    Dim Datos As ADODB.Recordset, Filas As ADODB.Recordset
    Dim Conexion As ADODB.Connection
    Dim Codigos As Collection
    Dim Fila As Long, Columna As Integer, Ultima As Long, UltimaColumna As Long
    Dim Encabezados As Variant
    Dim Valor As Variant
    Dim Texto As String
    Dim Documento As DatosDocumento
    Dim EnTransaccion As Boolean
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Len(Trim$(Plantilla.Descripcion)) = 0 Then Err.Raise vbObjectError + 310, , "Debe especificar la descripcion de la importacion."
    If Dir$(Ruta) = "" Then Err.Raise vbObjectError + 311, , "No existe el archivo de Excel."
    Set Distritos = CreateObject("Scripting.Dictionary")
    Distritos.CompareMode = vbTextCompare
    Set Datos = ConsultarDatos("SELECT CodPos,CoDist FROM dbo.MDBDistritos WHERE Active=1")
    Do Until Datos.EOF
        Distritos.Add Trim$(CStr(Datos.Fields(0).Value)), CLng(Datos.Fields(1).Value)
        Datos.MoveNext
    Loop
    Datos.Close
    Set Filas = New ADODB.Recordset
    Filas.Fields.Append "Empresa", adVarWChar, 500
    Filas.Fields.Append "Direccion", adVarWChar, 500
    Filas.Fields.Append "Distrito", adInteger
    Filas.Fields.Append "Contacto", adVarWChar, 150
    Filas.Open
    Set Excel = CreateObject("Excel.Application")
    Excel.Visible = False
    Excel.DisplayAlerts = False
    Excel.EnableEvents = False
    Excel.AutomationSecurity = 3
    Set Libro = Excel.Workbooks.Open(Ruta, 0, True)
    Set Hoja = Libro.ActiveSheet
    Encabezados = Array("EMPRESA", "DIRECCION", "DISTRITO", "CONTACTO")
    For Columna = 1 To 4
        Valor = Hoja.Cells(1, Columna).Value2
        If IsError(Valor) Then Err.Raise vbObjectError + 312, , "El encabezado contiene un error de Excel."
        If UCase$(Trim$(CStr(Valor))) <> Encabezados(Columna - 1) Then Err.Raise vbObjectError + 313, , "Las columnas deben ser EMPRESA, DIRECCION, DISTRITO y CONTACTO."
        UltimaColumna = Hoja.Cells(Hoja.Rows.Count, Columna).End(-4162).Row
        If UltimaColumna > Ultima Then Ultima = UltimaColumna
    Next
    If Ultima < 2 Then Err.Raise vbObjectError + 314, , "El archivo no contiene registros."
    If Ultima > 16000 Then Err.Raise vbObjectError + 315, , "El archivo supera los 15999 registros permitidos. Divida la informacion en varios archivos."
    For Fila = 2 To Ultima
        Filas.AddNew
        For Columna = 1 To 4
            Valor = Hoja.Cells(Fila, Columna).Value2
            If IsError(Valor) Then Err.Raise vbObjectError + 316, , "La celda contiene un error de Excel."
            Texto = UCase$(Trim$(CStr(Valor)))
            If InStr(Texto, vbCr) > 0 Or InStr(Texto, vbLf) > 0 Then Err.Raise vbObjectError + 317, , "La celda contiene saltos de linea."
            Select Case Columna
                Case 1, 2
                    If Len(Texto) = 0 Or Len(Texto) > 500 Then Err.Raise vbObjectError + 318, , "Empresa y direccion son obligatorias y admiten hasta 500 caracteres."
                    Filas.Fields(Columna - 1).Value = Texto
                Case 3
                    If Not Distritos.Exists(Texto) Then Err.Raise vbObjectError + 319, , "El codigo postal no existe en el catalogo de distritos."
                    Filas.Fields(2).Value = Distritos.Item(Texto)
                Case 4
                    If Len(Texto) > 150 Then Err.Raise vbObjectError + 320, , "El contacto supera los 150 caracteres."
                    Filas.Fields(3).Value = Texto
            End Select
        Next
        Filas.Update
    Next
    'Cerrar la instancia propia de Excel antes de bloquear registros en SQL.
    Set Hoja = Nothing
    Libro.Close False: Set Libro = Nothing
    Excel.Quit: Set Excel = Nothing
    Set Conexion = AbrirConexion()
    Conexion.BeginTrans: EnTransaccion = True
    Set Codigos = New Collection
    Filas.MoveFirst
    Fila = 2
    Do Until Filas.EOF
        Documento.Servicio = Plantilla.Servicio: Documento.Cliente = Plantilla.Cliente
        Documento.Producto = Plantilla.Producto: Documento.Referencia = Plantilla.Referencia
        Documento.Descripcion = Plantilla.Descripcion
        Documento.Destinatario = CStr(Filas.Fields(0).Value)
        Documento.Direccion = CStr(Filas.Fields(1).Value)
        Documento.Distrito = CLng(Filas.Fields(2).Value)
        Documento.Contacto = CStr(Filas.Fields(3).Value)
        Documento.CodigoDestinatario = "0"
        Codigos.Add GuardarRegistro(Documento, Conexion)
        Filas.MoveNext
        Fila = Fila + 1
    Loop
    Conexion.CommitTrans: EnTransaccion = False
    Conexion.Close: Filas.Close
    Set ImportarDocumentos = Codigos
    Exit Function
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    If Fila >= 2 Then Descripcion = "Fila " & CStr(Fila) & ": " & Descripcion
    On Error Resume Next
    Set Hoja = Nothing
    If Not Libro Is Nothing Then Libro.Close False
    Set Libro = Nothing
    If Not Excel Is Nothing Then Excel.Quit
    Set Excel = Nothing
    If Not Conexion Is Nothing Then
        If EnTransaccion Then Conexion.RollbackTrans
        Conexion.Close
    End If
    If Not Filas Is Nothing Then Filas.Close
    If Not Datos Is Nothing Then Datos.Close
    On Error GoTo 0
    Err.Raise Numero, "ImportarDocumentos", Descripcion
End Function
