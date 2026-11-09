Attribute VB_Name = "MdiExport"
Option Explicit

Public Sub CmdGenExcel(ByVal Cuadricula As MSFlexGridLib.MSFlexGrid, ByVal Titulo As String, Optional ByVal RutaArchivo As String = "", Optional ByVal Contexto As String = "")
    Dim Valores() As Variant, Fila As Long, Columna As Integer
    If Cuadricula.Rows < 2 Then Err.Raise 5, , "No existen registros para exportar."
    ReDim Valores(1 To Cuadricula.Rows, 1 To Cuadricula.Cols)
    For Fila = 0 To Cuadricula.Rows - 1
        For Columna = 0 To Cuadricula.Cols - 1: Valores(Fila + 1, Columna + 1) = Cuadricula.TextMatrix(Fila, Columna): Next
    Next
    GenerarLibro Valores, Cuadricula.Rows - 1, Cuadricula.Cols, Titulo, RutaArchivo, Contexto
End Sub

Public Sub ExportarDatos(ByVal Datos As ADODB.Recordset, ByVal Titulo As String, Optional ByVal RutaArchivo As String = "")
    Dim Copia As ADODB.Recordset, Valores() As Variant, Fila As Long, Columna As Integer
    If Datos.RecordCount < 1 Then Err.Raise 5, , "El informe no contiene registros."
    ReDim Valores(1 To Datos.RecordCount + 1, 1 To Datos.Fields.Count)
    For Columna = 0 To Datos.Fields.Count - 1: Valores(1, Columna + 1) = Datos.Fields(Columna).Name: Next
    Set Copia = Datos.Clone: Copia.MoveFirst: Fila = 2
    Do Until Copia.EOF
        For Columna = 0 To Copia.Fields.Count - 1: Valores(Fila, Columna + 1) = TextoCampo(Copia.Fields(Columna)): Next
        Fila = Fila + 1: Copia.MoveNext
    Loop
    Copia.Close
    GenerarLibro Valores, Datos.RecordCount, Datos.Fields.Count, Titulo, RutaArchivo
End Sub

Private Sub GenerarLibro(ByRef Valores() As Variant, ByVal Filas As Long, ByVal Columnas As Integer, ByVal Titulo As String, ByVal RutaArchivo As String, Optional ByVal Contexto As String = "")
    Dim Aplicacion As Object, Libro As Object, Hoja As Object
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Filas > 65530 Then Err.Raise 5, , "El formato XLS admite hasta 65530 registros con este encabezado. Ajuste los filtros."
    Set Aplicacion = CreateObject("Excel.Application"): Set Libro = Aplicacion.Workbooks.Add: Set Hoja = Libro.Worksheets(1)
    With Hoja.Range(Hoja.Cells(6, 1), Hoja.Cells(Filas + 6, Columnas))
        .NumberFormat = "@": .Value2 = Valores: .Font.Name = "Arial": .Font.Size = 8
    End With
    With Hoja.Range(Hoja.Cells(6, 1), Hoja.Cells(6, Columnas))
        .Font.Bold = True: .Font.ColorIndex = 2: .Interior.ColorIndex = 54
    End With
    Hoja.Cells(1, 1).Value2 = "REPORTE GENERADO POR CLEVER SOFTPRO"
    Hoja.Cells(2, 1).Value2 = Titulo
    Hoja.Cells(3, 1).Value2 = "FECHA DEL REPORTE: " & Format$(Date, "dd/mm/yyyy")
    Hoja.Cells(4, 1).Value2 = "REGISTROS ENCONTRADOS: " & CStr(Filas)
    Hoja.UsedRange.Columns.AutoFit
    If Len(Contexto) > 0 Then
        With Hoja.Range(Hoja.Cells(5, 1), Hoja.Cells(5, Columnas))
            .Merge: .Value2 = Contexto: .WrapText = True
            .Font.Name = "Arial": .Font.Size = 8
            .RowHeight = 36
        End With
    End If
    If RutaArchivo = "" Then
        Aplicacion.Visible = True: Aplicacion.UserControl = True
    Else
        Libro.SaveAs RutaArchivo, 56
        Libro.Close False: Aplicacion.Quit
    End If
    Set Hoja = Nothing: Set Libro = Nothing: Set Aplicacion = Nothing
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Libro Is Nothing Then Libro.Close False
    If Not Aplicacion Is Nothing Then Aplicacion.Quit
    Set Hoja = Nothing: Set Libro = Nothing: Set Aplicacion = Nothing
    On Error GoTo 0
    Err.Raise Numero, "GenerarLibro", Descripcion
End Sub
