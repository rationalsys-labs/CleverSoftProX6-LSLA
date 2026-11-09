Attribute VB_Name = "MdiEstados"
Option Explicit

Public Type DatosConfirmacion
    Codigo As String
    Estado As Long
    Fecha As Date
    Informe As String
End Type

Private Function TextoXml(ByVal Texto As String) As String
    TextoXml = Replace(Replace(Replace(Replace(Replace(Texto, "&", "&amp;"), "<", "&lt;"), ">", "&gt;"), """", "&quot;"), "'", "&apos;")
    TextoXml = Replace(Replace(Replace(TextoXml, vbCr, "&#xD;"), vbLf, "&#xA;"), vbTab, "&#x9;")
End Function

Public Sub GuardarConfirmaciones(ByRef Confirmaciones() As DatosConfirmacion, ByVal Cantidad As Long)
    Dim Conexion As ADODB.Connection, Comando As ADODB.Command
    Dim Lista As String, Posicion As Long
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Cantidad < 1 Then Err.Raise 5, , "No existen confirmaciones para registrar."
    Lista = "<Confirmaciones>"
    For Posicion = 1 To Cantidad
        With Confirmaciones(Posicion)
            If Not CodigoValido(.Codigo) Or .Estado <= 1 Or Len(.Informe) > 500 Then Err.Raise 5, , "La confirmacion contiene datos no validos."
            If .Fecha < DateSerial(1753, 1, 1) Then Err.Raise 5, , "La fecha no es valida para la base de datos."
            Lista = Lista & "<Confirmacion Codigo=""" & CStr(CDec(.Codigo)) & """ Estado=""" & CStr(.Estado) & """ Fecha=""" & Format$(.Fecha, "yyyy-mm-dd") & "T" & Format$(.Fecha, "hh:nn:ss") & """><Informe>" & TextoXml(.Informe) & "</Informe></Confirmacion>"
        End With
    Next
    Lista = Lista & "</Confirmaciones>"
    Set Conexion = AbrirConexion()
    Set Comando = New ADODB.Command
    Set Comando.ActiveConnection = Conexion
    Comando.CommandType = adCmdStoredProc: Comando.CommandText = "dbo.GuardarConfirmaciones"
    Comando.Parameters.Append Comando.CreateParameter("@Usuario", adInteger, adParamInput, , UsuarioOperativo())
    Comando.Parameters.Append Comando.CreateParameter("@Confirmaciones", adLongVarWChar, adParamInput, Len(Lista), Lista)
    Comando.Execute , , adExecuteNoRecords
    Conexion.Close
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Conexion Is Nothing Then Conexion.Close
    On Error GoTo 0
    Err.Raise Numero, "GuardarConfirmaciones", Descripcion
End Sub

Public Function ConsultarDelivery(ByVal Mensajero As Long, ByVal Desde As Date, ByVal Hasta As Date) As ADODB.Recordset
    Dim Conexion As ADODB.Connection, Comando As ADODB.Command, Datos As ADODB.Recordset
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Mensajero <= 0 Or DateValue(Desde) > DateValue(Hasta) Then Err.Raise 5, , "Seleccione un mensajero y un intervalo valido."
    Set Conexion = AbrirConexion()
    Set Comando = New ADODB.Command
    Set Comando.ActiveConnection = Conexion
    Comando.CommandType = adCmdStoredProc: Comando.CommandText = "dbo.ConsultarDelivery"
    Comando.Parameters.Append Comando.CreateParameter("@Mensajero", adInteger, adParamInput, , Mensajero)
    Comando.Parameters.Append Comando.CreateParameter("@Desde", adDBTimeStamp, adParamInput, , DateValue(Desde))
    Comando.Parameters.Append Comando.CreateParameter("@Hasta", adDBTimeStamp, adParamInput, , DateValue(Hasta))
    Set Datos = New ADODB.Recordset
    Datos.CursorLocation = adUseClient
    Datos.Open Comando, , adOpenStatic, adLockReadOnly
    Set Datos.ActiveConnection = Nothing
    Conexion.Close
    Set ConsultarDelivery = Datos
    Exit Function
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    If Not Conexion Is Nothing Then Conexion.Close
    On Error GoTo 0
    Err.Raise Numero, "ConsultarDelivery", Descripcion
End Function

Public Sub MostrarHistorial(ByVal Codigo As String, ByVal Arbol As Object, Optional ByVal EstadoPendiente As String = "", Optional ByVal FechaPendiente As Date = 0)
    Dim Datos As ADODB.Recordset, Pendiente As Boolean, Nodo As Object
    Arbol.Nodes.Clear
    Set Datos = ConsultarDatos("SELECT Status,Emision,Item,Usuario,Informe FROM dbo.ViewHistory WHERE Codigo=? ORDER BY Emision,Item", Codigo, True)
    Pendiente = (Len(EstadoPendiente) > 0)
    Do Until Datos.EOF
        If Pendiente And FechaPendiente < CDate(Datos.Fields("Emision").Value) Then
            Set Nodo = Arbol.Nodes.Add(, , "PENDIENTE", EstadoPendiente & " - " & Format$(FechaPendiente, "dd/mm/yyyy hh:nn:ss") & " (sin guardar)", MDIPri.ICO.ListImages(23).Key)
            Nodo.BackColor = vbRed: Pendiente = False
        End If
        Arbol.Nodes.Add , , "H" & CStr(Datos.Fields("Item").Value), TextoDato(Datos.Fields("Status").Value) & " - " & Format$(Datos.Fields("Emision").Value, "dd/mm/yyyy hh:nn:ss") & " " & TextoDato(Datos.Fields("Usuario").Value) & " " & TextoDato(Datos.Fields("Informe").Value), MDIPri.ICO.ListImages(23).Key
        Datos.MoveNext
    Loop
    Datos.Close
    If Pendiente Then
        Set Nodo = Arbol.Nodes.Add(, , "PENDIENTE", EstadoPendiente & " - " & Format$(FechaPendiente, "dd/mm/yyyy hh:nn:ss") & " (sin guardar)", MDIPri.ICO.ListImages(23).Key)
        Nodo.BackColor = vbRed
    End If
End Sub