Attribute VB_Name = "MdiImpresion"
Option Explicit

Public Function PatronBarras(ByVal Codigo As String) As String
    Dim Patrones As Variant
    Dim Posicion As Integer
    If Not CodigoValido(Codigo) Then Err.Raise vbObjectError + 330, , "Codigo no valido para impresion."
    'Patrones numericos del generador de barras del proyecto original.
    Patrones = Array("11 221", "21 112", "12 112", "22 111", "11 212", "21 211", "12 211", "11 122", "21 121", "12 121")
    PatronBarras = "1 1221"
    For Posicion = 1 To Len(Codigo)
        PatronBarras = PatronBarras & Patrones(CInt(Mid$(Codigo, Posicion, 1)))
    Next
    PatronBarras = PatronBarras & "1 1221"
End Function

Private Sub DibujarBarras(ByVal Codigo As String, ByVal Izquierda As Single, ByVal Arriba As Single)
    Dim Patron As String
    Dim Posicion As Integer
    Dim Ancho As Single, Inicio As Single
    Patron = PatronBarras(Codigo)
    Inicio = Izquierda + 2.5
    For Posicion = 1 To Len(Patron)
        Select Case Mid$(Patron, Posicion, 1)
            Case " "
                Inicio = Inicio + 0.25
            Case "1", "2"
                Inicio = Inicio + 0.25
                Ancho = CSng(Mid$(Patron, Posicion, 1)) * 0.25
                Printer.Line (Inicio, Arriba)-(Inicio + Ancho, Arriba + 8), vbBlack, BF
                Inicio = Inicio + Ancho
        End Select
    Next
End Sub

Public Function TextoRegistro(ByVal Codigo As String) As String
    Dim Datos As ADODB.Recordset
    Set Datos = ConsultarDatos("SELECT Codigo,Cliente,Empresa,Direccion,Distrito,Contacto,Descripcion,Fecha,Boleta,Producto,Documento FROM dbo.ViewDocument v JOIN dbo.MDBDocument d ON d.NroCod=v.Codigo AND d.Active=1 WHERE v.Codigo=?", Codigo, True)
    If Datos.EOF Then Datos.Close: Err.Raise vbObjectError + 331, , "No se encontro el documento para imprimir."
    TextoRegistro = "Codigo: " & CStr(Datos.Fields(0).Value) & vbCrLf & _
        "Cliente: " & TextoDato(Datos.Fields(1).Value) & vbCrLf & _
        "Destinatario: " & TextoDato(Datos.Fields(2).Value) & vbCrLf & _
        "Direccion: " & TextoDato(Datos.Fields(3).Value) & vbCrLf & _
        "Distrito: " & TextoDato(Datos.Fields(4).Value) & vbCrLf & _
        "Contacto: " & TextoDato(Datos.Fields(5).Value) & vbCrLf & _
        "Descripcion: " & TextoDato(Datos.Fields(6).Value) & vbCrLf & _
        "Fecha: " & Format$(Datos.Fields(7).Value, "dd/mm/yyyy hh:nn:ss") & vbCrLf & _
        "Boleta: " & TextoDato(Datos.Fields(8).Value) & vbCrLf & _
        "Producto: " & TextoDato(Datos.Fields(9).Value) & vbCrLf & _
        "Referencia: " & TextoDato(Datos.Fields(10).Value)
    Datos.Close
End Function

Public Sub ImprimirLinea(ByVal Texto As String, ByRef Arriba As Single, ByVal Ancho As Single)
    Dim Longitud As Long, Corte As Long
    Do While Len(Texto) > 0
        Longitud = Len(Texto)
        Do While Longitud > 1 And Printer.TextWidth(Left$(Texto, Longitud)) > Ancho
            Longitud = Longitud - 1
        Loop
        If Longitud < Len(Texto) Then
            Corte = InStrRev(Left$(Texto, Longitud), " ")
            If Corte > 1 Then Longitud = Corte
        End If
        If Arriba + 6 > Printer.ScaleHeight - 15 Then Printer.NewPage: Arriba = 15
        Printer.CurrentX = 15: Printer.CurrentY = Arriba
        Printer.Print Left$(Texto, Longitud)
        Texto = Mid$(Texto, Longitud + 1)
        Arriba = Arriba + 5
    Loop
End Sub

Public Sub ImprimirRegistros(ByVal Codigos As Collection, ByVal Modalidad As String)
    Dim Codigo As Variant, Lineas As Variant, Linea As Variant
    Dim Arriba As Single
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Codigos.Count = 0 Then Exit Sub
    Printer.ScaleMode = vbMillimeters
    Printer.FontName = "Arial": Printer.FontSize = 10
    Arriba = 15
    For Each Codigo In Codigos
        If Arriba + 90 > Printer.ScaleHeight - 15 Then Printer.NewPage: Arriba = 15
        ImprimirLinea "Clever SoftPro - Constancia de servicio", Arriba, Printer.ScaleWidth - 30
        DibujarBarras CStr(Codigo), 15, Arriba
        Arriba = Arriba + 10
        Lineas = Split(TextoRegistro(CStr(Codigo)), vbCrLf)
        For Each Linea In Lineas: ImprimirLinea CStr(Linea), Arriba, Printer.ScaleWidth - 30: Next
        If Modalidad = "Entrega y Servicio" Then
            ImprimirLinea "Constancia de entrega - Codigo " & CStr(Codigo), Arriba, Printer.ScaleWidth - 30
            ImprimirLinea "Recibido por: _____________________  Fecha: __________  Firma: ______________", Arriba, Printer.ScaleWidth - 30
        End If
        Arriba = Arriba + 8
    Next
    Printer.EndDoc
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    Printer.KillDoc
    On Error GoTo 0
    Err.Raise Numero, "ImprimirRegistros", Descripcion
End Sub

Public Function TextoCargo(ByVal Cuadricula As Object, ByVal Fila As Long, Optional ByVal FechaEsReparto As Boolean = False) As String
    Dim Rotulo As String
    If FechaEsReparto Then Rotulo = "FECHA REPARTO: " Else Rotulo = "FECHA DOCUMENTO: "
    TextoCargo = "PRODUCTO: " & Cuadricula.TextMatrix(Fila, 1) & vbCrLf & _
        Rotulo & Cuadricula.TextMatrix(Fila, 6) & "  CLIENTE: " & Cuadricula.TextMatrix(Fila, 8) & vbCrLf & _
        "DESTINATARIO: " & Cuadricula.TextMatrix(Fila, 2) & vbCrLf & _
        "DIRECCION: " & Cuadricula.TextMatrix(Fila, 3) & vbCrLf & _
        "DISTRITO: " & Cuadricula.TextMatrix(Fila, 4) & "  CONTACTO: " & Cuadricula.TextMatrix(Fila, 5) & vbCrLf & _
        "MENSAJERO: " & Cuadricula.TextMatrix(Fila, 7)
End Function

Private Sub CabeceraCargo(ByVal Cargo As String, ByVal Pagina As Long, ByRef Arriba As Single, Optional ByVal FechaReparto As Date = 0)
    Printer.CurrentX = 15: Printer.CurrentY = 12
    Printer.Print "REGISTRO DE DOCUMENTOS A REPARTO"
    Printer.CurrentX = 15
    Printer.Print "CONTROL DE UNIDADES DE DISTRIBUCION"
    Printer.CurrentX = 15
    Printer.Print "REPORTE N. " & Cargo & "   PAGINA: " & CStr(Pagina)
    Printer.CurrentX = 15
    If FechaReparto > 0 Then Printer.Print "FECHA REPARTO: " & Format$(FechaReparto, "dd/mm/yyyy hh:nn:ss")
    Printer.CurrentX = 15
    Printer.Print "IMPRESO: " & Format$(Now, "dd/mm/yyyy hh:nn:ss")
    Arriba = 38
End Sub

Private Sub LineaCargo(ByVal Texto As String, ByVal Cargo As String, ByRef Pagina As Long, ByRef Arriba As Single, Optional ByVal FechaReparto As Date = 0)
    Dim Longitud As Long, Corte As Long
    Do While Len(Texto) > 0
        Longitud = Len(Texto)
        Do While Longitud > 1 And Printer.TextWidth(Left$(Texto, Longitud)) > Printer.ScaleWidth - 30
            Longitud = Longitud - 1
        Loop
        If Longitud < Len(Texto) Then
            Corte = InStrRev(Left$(Texto, Longitud), " ")
            If Corte > 1 Then Longitud = Corte
        End If
        If Arriba + 5 > Printer.ScaleHeight - 15 Then
            Printer.NewPage: Pagina = Pagina + 1
            CabeceraCargo Cargo, Pagina, Arriba, FechaReparto
        End If
        Printer.CurrentX = 15: Printer.CurrentY = Arriba
        Printer.Print Left$(Texto, Longitud)
        Texto = Mid$(Texto, Longitud + 1): Arriba = Arriba + 4.5
    Loop
End Sub

Public Sub ImprimirCargo(ByVal Cuadricula As Object, ByVal Cargo As String, Optional ByVal FechaReparto As Date = 0, Optional ByVal FechasDeReparto As Boolean = False)
    Dim Fila As Long, Pagina As Long, Arriba As Single, Linea As Variant
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Cuadricula.Rows < 2 Or Not CodigoValido(Cargo) Then Err.Raise 5, , "No existe un cargo para imprimir."
    Printer.ScaleMode = vbMillimeters
    Printer.FontName = "Arial": Printer.FontSize = 8
    If Printer.ScaleWidth < 100 Or Printer.ScaleHeight < 100 Then Err.Raise 5, , "Seleccione una impresora con hoja normal."
    Pagina = 1
    CabeceraCargo Cargo, Pagina, Arriba, FechaReparto
    For Fila = 1 To Cuadricula.Rows - 1
        If Arriba + 45 > Printer.ScaleHeight - 15 Then
            Printer.NewPage: Pagina = Pagina + 1
            CabeceraCargo Cargo, Pagina, Arriba, FechaReparto
        End If
        DibujarBarras Cuadricula.TextMatrix(Fila, 0), 15, Arriba
        Arriba = Arriba + 9
        LineaCargo "CODIGO: " & Cuadricula.TextMatrix(Fila, 0), Cargo, Pagina, Arriba, FechaReparto
        For Each Linea In Split(TextoCargo(Cuadricula, Fila, FechasDeReparto), vbCrLf)
            LineaCargo CStr(Linea), Cargo, Pagina, Arriba, FechaReparto
        Next
        Arriba = Arriba + 3
    Next
    LineaCargo "System Clever SoftPro - Reporte basico de Servicios", Cargo, Pagina, Arriba, FechaReparto
    LineaCargo "TOTAL DE DOCUMENTOS: " & CStr(Cuadricula.Rows - 1), Cargo, Pagina, Arriba, FechaReparto
    Printer.EndDoc
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    Printer.KillDoc
    On Error GoTo 0
    Err.Raise Numero, "ImprimirCargo", Descripcion
End Sub
