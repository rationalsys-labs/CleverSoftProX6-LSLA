VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmDelivery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Control de Documentos a Delivery"
   ClientHeight    =   7785
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13875
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   13875
   Begin VB.CommandButton CmdLimpiar
      Caption         =   "Limpiar Consulta"
      Height          =   330
      Left            =   11400
      TabIndex        =   13
      Top             =   90
      Width           =   2340
   End
   Begin VB.ComboBox CmbMensajeros 
      Height          =   315
      Left            =   1125
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   90
      Width           =   3615
   End
   Begin VB.CommandButton CmdCargar 
      Caption         =   "Cargar"
      Height          =   330
      Left            =   9675
      TabIndex        =   10
      Top             =   90
      Width           =   1545
   End
   Begin MSComCtl2.DTPicker DtpDes 
      Height          =   285
      Left            =   5670
      TabIndex        =   7
      Top             =   90
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   48824321
      CurrentDate     =   40548
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estadistica"
      Height          =   7035
      Left            =   90
      TabIndex        =   2
      Top             =   405
      Width           =   3210
      Begin VB.CommandButton CmdGenExcel 
         Caption         =   "Generar Excel"
         Height          =   330
         Left            =   135
         TabIndex        =   12
         Top             =   6570
         Width           =   1770
      End
      Begin MSFlexGridLib.MSFlexGrid GridSta 
         Height          =   4425
         Left            =   90
         TabIndex        =   3
         Top             =   225
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   7805
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorFixed  =   8421504
         BackColorBkg    =   -2147483633
         GridLines       =   0
         GridLinesFixed  =   0
         BorderStyle     =   0
         Appearance      =   0
         FormatString    =   ""
      End
   End
   Begin MSFlexGridLib.MSFlexGrid GridReport 
      Height          =   5640
      Left            =   3375
      TabIndex        =   0
      Top             =   450
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   9948
      _Version        =   393216
      FixedCols       =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   4
      Top             =   7545
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   3704
            MinWidth        =   3704
            TextSave        =   "06/01/2011"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   19403
            MinWidth        =   19403
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker DtpHas 
      Height          =   285
      Left            =   8010
      TabIndex        =   9
      Top             =   90
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   503
      _Version        =   393216
      Format          =   48824321
      CurrentDate     =   40548
   End
   Begin MSFlexGridLib.MSFlexGrid GridAvance 
      Height          =   1275
      Left            =   3375
      TabIndex        =   11
      Top             =   6165
      Width           =   10365
      _ExtentX        =   18283
      _ExtentY        =   2249
      _Version        =   393216
   End
   Begin VB.Label Label3 
      Caption         =   "Hasta:"
      Height          =   195
      Left            =   7335
      TabIndex        =   8
      Top             =   135
      Width           =   600
   End
   Begin VB.Label Label2 
      Caption         =   "Desde:"
      Height          =   195
      Left            =   4995
      TabIndex        =   6
      Top             =   135
      Width           =   600
   End
   Begin VB.Label Label1 
      Caption         =   "Mensajeros:"
      Height          =   240
      Left            =   135
      TabIndex        =   5
      Top             =   135
      Width           =   825
   End
End
Attribute VB_Name = "FrmDelivery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Cargando As Boolean
Private Consultando As Boolean
Private ConsultaValida As Boolean
Private ContextoConsulta As String

Private Sub Form_Load()
    Dim Campos As Variant, Anchos As Variant, Columna As Integer
    Dim Panel As MSComctlLib.Panel
    On Error GoTo ErrInt
    Cargando = True
    For Each Panel In StatusBar1.Panels: Panel.MinWidth = 15: Panel.Style = sbrText: Next
    StatusBar1.Panels(1).Width = 2200: StatusBar1.Panels(2).AutoSize = sbrSpring: StatusBar1.Panels(3).Width = 3600
    Me.Icon = MDIPri.ICO.ListImages(25).Picture
    DtpHas.Value = Date: DtpDes.Value = Date
    GridSta.Rows = 1: GridSta.Cols = 2
    GridSta.TextMatrix(0, 0) = "Status": GridSta.ColWidth(0) = 2000
    GridSta.TextMatrix(0, 1) = "Cantidad": GridSta.ColWidth(1) = 800
    GridReport.Rows = 1: GridReport.Cols = 8
    Campos = Array("Codigo", "Descripcion", "Empresa", "Contacto", "Emision", "Informe", "Status", "Fecha")
    Anchos = Array(900, 1600, 2000, 1500, 1200, 0, 1200, 1500)
    For Columna = 0 To 7
        GridReport.TextMatrix(0, Columna) = Campos(Columna): GridReport.ColWidth(Columna) = Anchos(Columna)
    Next
    GridAvance.Rows = 1: GridAvance.Cols = 6
    Campos = Array("Mensajero", "Entregados", "Devoluciones", "Pendientes", "Total", "Registros")
    For Columna = 0 To 5
        GridAvance.TextMatrix(0, Columna) = Campos(Columna): GridAvance.ColWidth(Columna) = 1200
    Next
    GridAvance.ColWidth(0) = 3800
    CargarLista CmbMensajeros, "SELECT CodMen,NomMen FROM dbo.Mdbmensajeros WHERE Active=1 ORDER BY NomMen"
    Cargando = False: ActualizarEstado
    Exit Sub
ErrInt:
    Cargando = False
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub LimpiarConsulta()
    GridReport.Rows = 1: GridSta.Rows = 1: GridAvance.Rows = 1
    ConsultaValida = False: ContextoConsulta = ""
    ActualizarEstado
End Sub

Public Sub CargarDelivery()
    Dim Datos As ADODB.Recordset, Fila As Long, Columna As Integer, Estado As Long
    Dim Entregados As Long, Pendientes As Long, Total As Long
    Dim Cantidades As Object, Nombres As Object, Clave As Variant
    On Error GoTo ErrInt
    If Consultando Then Exit Sub
    LimpiarConsulta
    If CmbMensajeros.ListIndex < 0 Then Err.Raise 5, , "Seleccione el mensajero que desea consultar."
    Consultando = True: ActualizarEstado
    Set Datos = MdiEstados.ConsultarDelivery(CmbMensajeros.ItemData(CmbMensajeros.ListIndex), DtpDes.Value, DtpHas.Value)
    Set Cantidades = CreateObject("Scripting.Dictionary"): Set Nombres = CreateObject("Scripting.Dictionary")
    Do Until Datos.EOF
        GridReport.Rows = GridReport.Rows + 1: Fila = GridReport.Rows - 1
        For Columna = 0 To 7: GridReport.TextMatrix(Fila, Columna) = TextoDato(Datos.Fields(Columna).Value): Next
        GridReport.TextMatrix(Fila, 4) = Format$(Datos.Fields("Emision").Value, "dd/mm/yyyy")
        If Not IsNull(Datos.Fields("Fecha").Value) Then GridReport.TextMatrix(Fila, 7) = Format$(Datos.Fields("Fecha").Value, "dd/mm/yyyy hh:nn:ss")
        Estado = CLng(Datos.Fields("CodEstado").Value)
        If Not Cantidades.Exists(Estado) Then Cantidades.Add Estado, 0: Nombres.Add Estado, TextoDato(Datos.Fields("Status").Value)
        Cantidades(Estado) = CLng(Cantidades(Estado)) + 1
        If Estado = 1 Then Pendientes = Pendientes + 1
        If Estado = 2 Then Entregados = Entregados + 1
        Total = Total + 1
        Datos.MoveNext
    Loop
    Datos.Close
    For Each Clave In Cantidades.Keys
        GridSta.Rows = GridSta.Rows + 1: Fila = GridSta.Rows - 1
        GridSta.TextMatrix(Fila, 0) = CStr(Nombres(Clave)): GridSta.TextMatrix(Fila, 1) = Format$(Cantidades(Clave), "00")
    Next
    If Total > 0 Then
        GridAvance.Rows = 2
        GridAvance.TextMatrix(1, 0) = CmbMensajeros.Text
        GridAvance.TextMatrix(1, 1) = Format$(Entregados, "00")
        GridAvance.TextMatrix(1, 2) = Format$(Total - Entregados - Pendientes, "00")
        GridAvance.TextMatrix(1, 3) = Format$(Pendientes, "00")
        GridAvance.TextMatrix(1, 4) = Format$(Total, "00"): GridAvance.TextMatrix(1, 5) = Format$(Total, "00")
    End If
    ContextoConsulta = "Mensajero: " & CmbMensajeros.Text & " (codigo " & CStr(CmbMensajeros.ItemData(CmbMensajeros.ListIndex)) & ") | Desde: " & Format$(DtpDes.Value, "dd/mm/yyyy") & " | Hasta: " & Format$(DtpHas.Value, "dd/mm/yyyy") & " | Estados actuales de las guias; intervalo por fecha de reparto."
    ConsultaValida = True: Consultando = False: ActualizarEstado
    Exit Sub
ErrInt:
    Dim Numero As Long, Descripcion As String
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    Consultando = False: LimpiarConsulta
    On Error GoTo 0
    Err.Raise Numero, "CargarDelivery", Descripcion
End Sub

Public Sub ExportarAvance(Optional ByVal RutaArchivo As String = "")
    If Consultando Or Not ConsultaValida Or GridReport.Rows < 2 Then Err.Raise 5, , "No existen registros para exportar."
    MdiExport.CmdGenExcel GridReport, "REPORTE DE AVANCE DE MENSAJEROS", RutaArchivo, ContextoConsulta
End Sub

Private Sub CmdCargar_Click()
    On Error GoTo ErrInt
    CargarDelivery
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdGenExcel_Click()
    On Error GoTo ErrInt
    ExportarAvance
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmbMensajeros_Click()
    If Not Cargando And Not Consultando Then LimpiarConsulta
End Sub

Private Sub DtpDes_Change()
    If Not Cargando And Not Consultando Then LimpiarConsulta
End Sub

Private Sub DtpHas_Change()
    If Not Cargando And Not Consultando Then LimpiarConsulta
End Sub
Private Sub ActualizarEstado()
    CmdCargar.Enabled = (Not Cargando And Not Consultando And CmbMensajeros.ListIndex >= 0 And DateValue(DtpDes.Value) <= DateValue(DtpHas.Value))
    CmdGenExcel.Enabled = (ConsultaValida And GridReport.Rows > 1 And Not Consultando)
    CmdLimpiar.Enabled = Not Consultando
    CmbMensajeros.Enabled = Not Consultando: DtpDes.Enabled = Not Consultando: DtpHas.Enabled = Not Consultando
    If Consultando Then
        StatusBar1.Panels(1).Text = "Consultando delivery"
    ElseIf ConsultaValida Then
        StatusBar1.Panels(1).Text = CStr(GridReport.Rows - 1) & " registros consultados"
    ElseIf DateValue(DtpDes.Value) > DateValue(DtpHas.Value) Then
        StatusBar1.Panels(1).Text = "Revise el intervalo de fechas"
    Else
        StatusBar1.Panels(1).Text = "Seleccione filtros y Cargar"
    End If
    If CmbMensajeros.ListIndex >= 0 Then StatusBar1.Panels(2).Text = CmbMensajeros.Text Else StatusBar1.Panels(2).Text = "Seleccione un mensajero"
    StatusBar1.Panels(3).Text = "Reparto: " & Format$(DtpDes.Value, "dd/mm/yyyy") & " - " & Format$(DtpHas.Value, "dd/mm/yyyy")
End Sub

Private Sub CmdLimpiar_Click()
    If Consultando Then Exit Sub
    Cargando = True
    CmbMensajeros.ListIndex = -1: DtpDes.Value = Date: DtpHas.Value = Date
    Cargando = False: LimpiarConsulta
End Sub
Private Sub Form_Activate()
    VerificarAccesoFormulario Me, 8
End Sub
