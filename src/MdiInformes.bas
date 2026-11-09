Attribute VB_Name = "MdiInformes"
Option Explicit

Public Type FiltroInforme
    Campo As String
    Operador As Integer
    Valor As String
    OpcionO As Boolean
End Type

Public Function ConsultaOrigen(ByVal Origen As Integer) As String
    Select Case Origen
        Case 1: ConsultaOrigen = "SELECT * FROM dbo.ViewInformeStatus"
        Case 2: ConsultaOrigen = "SELECT s.CodSer Codigo,s.NroBol Boleta,s.NomCli Cliente,s.CanSer Cantidad,p.DesPro Producto,t.DesTip Reparto,s.Grupo,CASE WHEN s.Estado=1 THEN 'Activo' ELSE 'Finalizado' END Estado FROM dbo.ViewDatosServicios s JOIN dbo.MDBProduct p ON p.CodPro=s.Tipo JOIN dbo.IntTipService t ON t.CodTip=s.TipSer"
        Case 3: ConsultaOrigen = "SELECT * FROM dbo.ViewCliente"
        Case 4: ConsultaOrigen = "SELECT i.Codigo,i.Emision,i.Producto,i.Docmto,i.Servicio,i.Boleta,i.Cliente,i.Empresa,i.Direccion,i.Distrito,i.Contacto,d.Descrip Descripcion FROM dbo.ViewInformeStatus i JOIN dbo.MDBDocument d ON d.NroCod=i.Codigo"
        Case Else: Err.Raise 5, , "Seleccione el origen del informe."
    End Select
End Function

Private Function CampoPermitido(ByVal Datos As ADODB.Recordset, ByVal Nombre As String) As ADODB.Field
    Dim Campo As ADODB.Field
    For Each Campo In Datos.Fields
        If StrComp(Campo.Name, Nombre, vbTextCompare) = 0 Then Set CampoPermitido = Campo: Exit Function
    Next
    Err.Raise 5, , "El campo del informe no pertenece al origen seleccionado."
End Function

Private Function EsDigitos(ByVal Texto As String) As Boolean
    Dim Posicion As Integer
    If Len(Texto) = 0 Then Exit Function
    For Posicion = 1 To Len(Texto)
        If Mid$(Texto, Posicion, 1) < "0" Or Mid$(Texto, Posicion, 1) > "9" Then Exit Function
    Next
    EsDigitos = True
End Function

Private Function FechaFiltro(ByVal Texto As String, ByRef SoloDia As Boolean) As Date
    Dim Partes As Variant, Fecha As Variant, Hora As Variant
    Dim Dia As Integer, Mes As Integer, Anio As Integer, Horas As Integer, Minutos As Integer, Segundos As Integer
    If Len(Trim$(Texto)) = 0 Then Err.Raise 5, , "Indique una fecha para el filtro. Use dd/mm/aaaa o aaaa-mm-dd."
    Partes = Split(Replace(Trim$(Texto), "T", " "), " ")
    If UBound(Partes) > 1 Then Err.Raise 5, , "Use fecha dd/mm/aaaa o aaaa-mm-dd y hora hh:mm:ss opcional."
    If InStr(Partes(0), "-") > 0 Then
        Fecha = Split(Partes(0), "-")
        If UBound(Fecha) <> 2 Then Err.Raise 5, , "Fecha no valida."
        If Not EsDigitos(Fecha(0)) Or Not EsDigitos(Fecha(1)) Or Not EsDigitos(Fecha(2)) Then Err.Raise 5, , "Fecha no valida."
        Anio = CInt(Fecha(0)): Mes = CInt(Fecha(1)): Dia = CInt(Fecha(2))
    Else
        Fecha = Split(Partes(0), "/")
        If UBound(Fecha) <> 2 Then Err.Raise 5, , "Use fecha dd/mm/aaaa o aaaa-mm-dd."
        If Not EsDigitos(Fecha(0)) Or Not EsDigitos(Fecha(1)) Or Not EsDigitos(Fecha(2)) Then Err.Raise 5, , "Fecha no valida."
        Dia = CInt(Fecha(0)): Mes = CInt(Fecha(1)): Anio = CInt(Fecha(2))
    End If
    If Anio < 1753 Or Anio > 9999 Or Mes < 1 Or Mes > 12 Or Dia < 1 Or Dia > 31 Then Err.Raise 5, , "Fecha no valida."
    FechaFiltro = DateSerial(Anio, Mes, Dia)
    If Year(FechaFiltro) <> Anio Or Month(FechaFiltro) <> Mes Or Day(FechaFiltro) <> Dia Then Err.Raise 5, , "Fecha no valida."
    SoloDia = (UBound(Partes) = 0)
    If Not SoloDia Then
        Hora = Split(Partes(1), ":")
        If UBound(Hora) <> 2 Then Err.Raise 5, , "Use hora hh:mm:ss."
        If Not EsDigitos(Hora(0)) Or Not EsDigitos(Hora(1)) Or Not EsDigitos(Hora(2)) Then Err.Raise 5, , "Hora no valida."
        Horas = CInt(Hora(0)): Minutos = CInt(Hora(1)): Segundos = CInt(Hora(2))
        If Horas > 23 Or Minutos > 59 Or Segundos > 59 Then Err.Raise 5, , "Hora no valida."
        FechaFiltro = FechaFiltro + TimeSerial(Horas, Minutos, Segundos)
    End If
End Function

Private Sub ParametroFecha(ByVal Comando As ADODB.Command, ByVal Fecha As Date)
    Comando.Parameters.Append Comando.CreateParameter("F" & CStr(Comando.Parameters.Count), adDBTimeStamp, adParamInput, , Fecha)
End Sub

Private Function CondicionFiltro(ByVal Comando As ADODB.Command, ByVal Campo As ADODB.Field, ByRef Filtro As FiltroInforme) As String
    Dim Nombre As String, Operador As String, Texto As String, Fecha As Date, Limite As Date, SoloDia As Boolean
    Dim Parametro As ADODB.Parameter
    Nombre = "[" & Campo.Name & "]"
    If Filtro.Operador < 1 Or Filtro.Operador > 12 Then Err.Raise 5, , "Seleccione una condicion valida."
    If Filtro.Operador >= 7 Then
        Texto = Replace(Replace(Replace(Filtro.Valor, "[", "[[]"), "%", "[%]"), "_", "[_]")
        Select Case Filtro.Operador
            Case 7, 8: Texto = Texto & "%"
            Case 9, 10: Texto = "%" & Texto
            Case 11, 12: Texto = "%" & Texto & "%"
        End Select
        If Len(Texto) > 4000 Then Err.Raise 5, , "El filtro es demasiado largo."
        If Filtro.Operador Mod 2 = 0 Then Operador = " NOT LIKE " Else Operador = " LIKE "
        Comando.Parameters.Append Comando.CreateParameter("T" & CStr(Comando.Parameters.Count), adVarWChar, adParamInput, Len(Texto) + 1, Texto)
        CondicionFiltro = "CONVERT(nvarchar(4000)," & Nombre & ")" & Operador & "?"
        Exit Function
    End If
    Operador = Choose(Filtro.Operador, "=", "<>", ">", ">=", "<", "<=")
    Select Case Campo.Type
        Case adDate, adDBTimeStamp, adDBDate
            Fecha = FechaFiltro(Filtro.Valor, SoloDia)
            If SoloDia Then
                If Fecha = DateSerial(9999, 12, 31) Then
                    Select Case Filtro.Operador
                        Case 1, 4
                            ParametroFecha Comando, Fecha: CondicionFiltro = Nombre & ">=?"
                        Case 2, 5
                            ParametroFecha Comando, Fecha: CondicionFiltro = Nombre & "<?"
                        Case 3: CondicionFiltro = "1=0"
                        Case 6: CondicionFiltro = Nombre & " IS NOT NULL"
                    End Select
                    Exit Function
                End If
                Limite = DateAdd("d", 1, Fecha)
                Select Case Filtro.Operador
                    Case 1, 2
                        ParametroFecha Comando, Fecha: ParametroFecha Comando, Limite
                        If Filtro.Operador = 1 Then
                            CondicionFiltro = "(" & Nombre & ">=? AND " & Nombre & "<?)"
                        Else
                            CondicionFiltro = "(" & Nombre & "<? OR " & Nombre & ">=?)"
                        End If
                        Exit Function
                    Case 3: Fecha = Limite: Operador = ">="
                    Case 6: Fecha = Limite: Operador = "<"
                End Select
            End If
            ParametroFecha Comando, Fecha
        Case adInteger, adSmallInt, adTinyInt, adUnsignedTinyInt, adBigInt, adNumeric, adDecimal, adCurrency, adDouble, adSingle
            If Not IsNumeric(Filtro.Valor) Then Err.Raise 5, , "El filtro numerico no es valido."
            Set Parametro = Comando.CreateParameter("N" & CStr(Comando.Parameters.Count), adNumeric, adParamInput)
            Parametro.Precision = 28: Parametro.NumericScale = 6: Parametro.Value = CDec(Filtro.Valor)
            Comando.Parameters.Append Parametro
        Case Else
            If Len(Filtro.Valor) > 4000 Then Err.Raise 5, , "El filtro es demasiado largo."
            Comando.Parameters.Append Comando.CreateParameter("T" & CStr(Comando.Parameters.Count), adVarWChar, adParamInput, Len(Filtro.Valor) + 1, Filtro.Valor)
    End Select
    CondicionFiltro = Nombre & Operador & "?"
End Function

Public Function ConsultarInforme(ByVal Origen As Integer, ByVal Columnas As Collection, ByRef Filtros() As FiltroInforme, ByVal Cantidad As Integer) As ADODB.Recordset
    Dim Conexion As ADODB.Connection, Comando As ADODB.Command, Metadatos As ADODB.Recordset, Datos As ADODB.Recordset
    Dim Campo As ADODB.Field, Columna As Variant, Seleccion As String, Condicion As String, Siguiente As String, Base As String
    Dim Posicion As Integer, Numero As Long, Descripcion As String, Nombres As Object
    On Error GoTo ErrInt
    If Columnas.Count = 0 Then Err.Raise 5, , "Seleccione las columnas del informe."
    If Cantidad < 0 Or Cantidad > 4 Then Err.Raise 5, , "El informe admite hasta cuatro filtros."
    Base = ConsultaOrigen(Origen)
    Set Metadatos = ConsultarDatos("SELECT TOP(0) * FROM (" & Base & ") Datos")
    Set Nombres = CreateObject("Scripting.Dictionary"): Nombres.CompareMode = vbTextCompare
    For Each Columna In Columnas
        Set Campo = CampoPermitido(Metadatos, CStr(Columna))
        If Nombres.Exists(Campo.Name) Then Err.Raise 5, , "No repita columnas en el informe."
        Nombres.Add Campo.Name, True
        If Seleccion <> "" Then Seleccion = Seleccion & ","
        Seleccion = Seleccion & "[" & Campo.Name & "]"
    Next
    Set Conexion = AbrirConexion(): Set Comando = New ADODB.Command
    Set Comando.ActiveConnection = Conexion: Comando.CommandType = adCmdText
    For Posicion = 0 To Cantidad - 1
        Set Campo = CampoPermitido(Metadatos, Filtros(Posicion).Campo)
        Siguiente = CondicionFiltro(Comando, Campo, Filtros(Posicion))
        If Posicion = 0 Then
            Condicion = Siguiente
        Else
            If Filtros(Posicion - 1).OpcionO Then Condicion = "(" & Condicion & " OR " & Siguiente & ")" Else Condicion = "(" & Condicion & " AND " & Siguiente & ")"
        End If
    Next
    Metadatos.Close
    Comando.CommandText = "SELECT " & Seleccion & " FROM (" & Base & ") Datos"
    If Condicion <> "" Then Comando.CommandText = Comando.CommandText & " WHERE " & Condicion
    Comando.CommandText = Comando.CommandText & " ORDER BY Codigo"
    Set Datos = New ADODB.Recordset: Datos.CursorLocation = adUseClient
    Datos.Open Comando, , adOpenStatic, adLockReadOnly
    Set Datos.ActiveConnection = Nothing: Conexion.Close
    Set ConsultarInforme = Datos
    Exit Function
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Metadatos Is Nothing Then Metadatos.Close
    If Not Datos Is Nothing Then Datos.Close
    If Not Conexion Is Nothing Then Conexion.Close
    Err.Clear
    On Error GoTo 0
    Err.Raise Numero, "ConsultarInforme", Descripcion
End Function



Public Function TextoCampo(ByVal Campo As ADODB.Field) As String
    If IsNull(Campo.Value) Then Exit Function
    If Campo.Type = adDate Or Campo.Type = adDBTimeStamp Or Campo.Type = adDBDate Then
        TextoCampo = Format$(Campo.Value, "dd/mm/yyyy hh:nn:ss")
    Else
        TextoCampo = CStr(Campo.Value)
    End If
End Function

Public Sub GuardarInformeTexto(ByVal Datos As ADODB.Recordset, ByVal Ruta As String, Optional ByVal Sobrescribir As Boolean = False)
    Dim Copia As ADODB.Recordset, Archivo As Object, Archivos As Object, Temporal As String, Linea As String, Campo As ADODB.Field
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Datos.EOF And Datos.BOF Then Err.Raise 5, , "El informe no contiene registros."
    Temporal = TemporalDestino(Ruta, Sobrescribir)
    Set Archivos = CreateObject("Scripting.FileSystemObject"): Set Archivo = Archivos.CreateTextFile(Temporal, False, True)
    For Each Campo In Datos.Fields
        If Linea <> "" Then Linea = Linea & vbTab
        Linea = Linea & CeldaTexto(Campo.Name)
    Next
    Archivo.WriteLine Linea
    Set Copia = Datos.Clone: Copia.MoveFirst
    Do Until Copia.EOF
        Linea = ""
        For Each Campo In Copia.Fields
            If Linea <> "" Then Linea = Linea & vbTab
            Linea = Linea & CeldaTexto(TextoCampo(Campo))
        Next
        Archivo.WriteLine Linea: Copia.MoveNext
    Loop
    Copia.Close: Set Copia = Nothing: Archivo.Close: Set Archivo = Nothing
    PublicarArchivo Temporal, Ruta, Sobrescribir: Temporal = ""
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Archivo Is Nothing Then Archivo.Close
    If Not Copia Is Nothing Then Copia.Close
    QuitarTemporal Temporal
    Err.Clear
    On Error GoTo 0
    Err.Raise Numero, "GuardarInformeTexto", Descripcion
End Sub

Public Sub ImprimirInforme(ByVal Datos As ADODB.Recordset, ByVal Titulo As String)
    Dim Copia As ADODB.Recordset, Campo As ADODB.Field, Arriba As Single
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Datos.RecordCount = 0 Then Err.Raise 5, , "El informe no contiene registros."
    Set Copia = Datos.Clone: Copia.MoveFirst
    Printer.ScaleMode = vbMillimeters: Printer.FontName = "Arial": Printer.FontSize = 9
    Arriba = 15
    MdiImpresion.ImprimirLinea Titulo, Arriba, Printer.ScaleWidth - 30
    MdiImpresion.ImprimirLinea "Fecha: " & Format$(Now, "dd/mm/yyyy hh:nn:ss"), Arriba, Printer.ScaleWidth - 30
    Do Until Copia.EOF
        For Each Campo In Copia.Fields
            MdiImpresion.ImprimirLinea Campo.Name & ": " & TextoCampo(Campo), Arriba, Printer.ScaleWidth - 30
        Next
        Arriba = Arriba + 5: Copia.MoveNext
    Loop
    Copia.Close: Printer.EndDoc
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Copia Is Nothing Then Copia.Close
    Printer.KillDoc
    Err.Clear
    On Error GoTo 0
    Err.Raise Numero, "ImprimirInforme", Descripcion
End Sub
