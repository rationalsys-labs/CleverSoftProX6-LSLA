VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FrmgestServ 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion y Control de Servicios por Pedido de Cliente"
   ClientHeight    =   8850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   18075
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8850
   ScaleWidth      =   18075
   Begin VB.ComboBox CmbBuscar 
      Height          =   315
      Left            =   8415
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   90
      Width           =   2130
   End
   Begin VB.TextBox txtBuscar 
      Height          =   285
      Left            =   4815
      TabIndex        =   19
      Top             =   90
      Width           =   3525
   End
   Begin VB.CommandButton CmdGenExcel 
      Caption         =   "Generar Documento Excel"
      Height          =   420
      Left            =   135
      TabIndex        =   17
      Top             =   8145
      Width           =   3255
   End
   Begin VB.CommandButton CmdMostrar 
      Caption         =   "Cargar Datos del Grupo"
      Height          =   420
      Left            =   135
      TabIndex        =   14
      Top             =   4725
      Width           =   3255
   End
   Begin VB.Frame Frame3 
      Caption         =   "Servicios Agrupados"
      Height          =   3390
      Left            =   135
      TabIndex        =   4
      Top             =   1260
      Width           =   3255
      Begin MSComctlLib.TreeView ViewAgrupar 
         Height          =   2130
         Left            =   135
         TabIndex        =   13
         Top             =   585
         Visible         =   0   'False
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   4075
         _Version        =   393217
         Indentation     =   19
         Style           =   7
         Checkboxes      =   -1  'True
         Appearance      =   1
      End
      Begin VB.Label LblRetirar
         Caption         =   "Suprimir: retirar servicio del grupo"
         Height          =   195
         Left            =   135
         Top             =   2730
         Width           =   3030
         TabIndex        =   24
      End
      Begin VB.CommandButton CmdAgru 
         Caption         =   "Agrupar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1980
         TabIndex        =   12
         Top             =   2970
         Width           =   1140
      End
      Begin VB.CommandButton CmdCanAgru 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   765
         TabIndex        =   11
         Top             =   2970
         Width           =   1140
      End
      Begin VB.CommandButton CmdAgrupar 
         Caption         =   "Agrupar Nuevos Servicios"
         Height          =   330
         Left            =   135
         TabIndex        =   10
         Top             =   225
         Width           =   3030
      End
      Begin MSFlexGridLib.MSFlexGrid GridAgrupar 
         Height          =   2130
         Left            =   135
         TabIndex        =   9
         Top             =   585
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   4154
         _Version        =   393216
         FixedCols       =   0
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Grupo"
      Height          =   1185
      Left            =   135
      TabIndex        =   3
      Top             =   45
      Width           =   3255
      Begin VB.CommandButton CmdNewGrup 
         Height          =   375
         Left            =   2655
         Picture         =   "FrmgestServ.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   23
         Top             =   225
         Width           =   465
      End
      Begin VB.CommandButton CmdBuscar 
         Height          =   375
         Left            =   2115
         Picture         =   "FrmgestServ.frx":058A
         Style           =   1  'Graphical
         TabIndex        =   16
         Top             =   225
         Width           =   465
      End
      Begin VB.Label LblNom 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   990
         TabIndex        =   8
         Top             =   675
         Width           =   2130
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   7
         Top             =   675
         Width           =   705
      End
      Begin VB.Label LblBol 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   990
         TabIndex        =   6
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estadisticas"
      Height          =   2895
      Left            =   135
      TabIndex        =   1
      Top             =   5175
      Width           =   3255
      Begin MSFlexGridLib.MSFlexGrid GridSta 
         Height          =   2535
         Left            =   90
         TabIndex        =   15
         Top             =   225
         Width           =   3075
         _ExtentX        =   5424
         _ExtentY        =   4471
         _Version        =   393216
         FixedCols       =   0
         BackColor       =   -2147483633
         BackColorFixed  =   8421504
         BackColorBkg    =   -2147483633
         GridLines       =   0
         GridLinesFixed  =   0
         BorderStyle     =   0
         Appearance      =   0
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   8115
      Left            =   3510
      TabIndex        =   0
      Top             =   450
      Width           =   14460
      _ExtentX        =   25506
      _ExtentY        =   14314
      _Version        =   393216
      FixedCols       =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   2
      Top             =   8610
      Width           =   18075
      _ExtentX        =   31882
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
            MinWidth        =   3704
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   9701
            MinWidth        =   9701
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   19403
            MinWidth        =   19403
         EndProperty
      EndProperty
   End
   Begin VB.Label lblCount 
      Alignment       =   1  'Right Justify
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   17100
      TabIndex        =   22
      Top             =   135
      Width           =   825
   End
   Begin VB.Label Label2 
      Caption         =   "Total de Registros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   15345
      TabIndex        =   21
      Top             =   135
      Width           =   1590
   End
   Begin VB.Label Label1 
      Caption         =   "Buscar Registro"
      Height          =   240
      Left            =   3510
      TabIndex        =   18
      Top             =   135
      Width           =   1275
   End
End
Attribute VB_Name = "FrmgestServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private RsRegist As ADODB.Recordset
Private CodigoGrupo As Long
Private Orden As Integer
Private Cargando As Boolean
Private Seleccionando As Boolean
Private Guardando As Boolean
Private ConsultaValida As Boolean
Private ConsultaPendiente As Boolean

Private Sub Form_Load()
    Dim Campos As Variant, Anchos As Variant, Columna As Integer
    Dim Panel As MSComctlLib.Panel
    On Error GoTo ErrInt
    Cargando = True
    Me.Icon = MDIPri.ICO.ListImages(25).Picture
    For Each Panel In StatusBar1.Panels
        Panel.MinWidth = 15
    Next
    StatusBar1.Panels(1).Width = 3000
    StatusBar1.Panels(2).AutoSize = sbrSpring
    StatusBar1.Panels(3).Width = 3900
    Orden = 1
    Campos = Array("Codigo", "Producto", "Empresa", "Direccion", "Distrito", "Contacto", "Descripcion", "Servicio", "Boleta", "Cliente", "UltimoSta", "Fecha", "Informe")
    Anchos = Array(1000, 0, 2500, 2500, 1500, 1800, 1500, 0, 0, 0, 1400, 0, 1800)
    Grid.Rows = 1: Grid.Cols = 13
    For Columna = 0 To 12
        Grid.TextMatrix(0, Columna) = Campos(Columna): Grid.ColWidth(Columna) = Anchos(Columna)
    Next
    Campos = Array("Codigo", "Empresa", "Direccion", "Distrito", "Contacto", "Servicio", "Boleta", "UltimoSta", "Informe")
    For Columna = 0 To UBound(Campos): CmbBuscar.AddItem Campos(Columna): Next
    CmbBuscar.ListIndex = 1: txtBuscar.MaxLength = 500
    GridAgrupar.Rows = 1: GridAgrupar.Cols = 3
    GridAgrupar.TextMatrix(0, 0) = "Codigo": GridAgrupar.ColWidth(0) = 600
    GridAgrupar.TextMatrix(0, 1) = "Boleta": GridAgrupar.ColWidth(1) = 500
    GridAgrupar.TextMatrix(0, 2) = "Cliente": GridAgrupar.ColWidth(2) = 1500
    GridSta.Rows = 1: GridSta.Cols = 2
    GridSta.TextMatrix(0, 0) = "Status": GridSta.ColWidth(0) = 2000
    GridSta.TextMatrix(0, 1) = "Cantidad": GridSta.ColWidth(1) = 800
    Cargando = False
    ActualizarEstado
    Exit Sub
ErrInt:
    Cargando = False
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Function ServiciosMarcados() As Long
    Dim Nodo As MSComctlLib.Node
    For Each Nodo In ViewAgrupar.Nodes
        If Nodo.Checked Then ServiciosMarcados = ServiciosMarcados + 1
    Next
End Function

Private Sub ActualizarEstado()
    Dim Consulta As Boolean, Marcados As Long
    If Cargando Then Exit Sub
    Consulta = (CodigoGrupo > 0 And Not Seleccionando And Not Guardando)
    Marcados = ServiciosMarcados()
    CmdBuscar.Enabled = Not Guardando: CmdNewGrup.Enabled = Not Guardando
    CmdMostrar.Enabled = Consulta
    CmdAgrupar.Enabled = (Consulta And Not ConsultaPendiente)
    CmdAgru.Enabled = (Seleccionando And Marcados > 0 And Not Guardando)
    CmdCanAgru.Enabled = (Seleccionando And Not Guardando)
    GridAgrupar.Enabled = (Consulta And Not ConsultaPendiente)
    ViewAgrupar.Enabled = Not Guardando
    CmbBuscar.Enabled = (Consulta And ConsultaValida And Not ConsultaPendiente)
    txtBuscar.Enabled = CmbBuscar.Enabled
    CmdGenExcel.Enabled = (CmbBuscar.Enabled And Grid.Rows > 1)
    If Guardando Then
        StatusBar1.Panels(1).Text = "Guardando grupo"
    ElseIf Seleccionando Then
        StatusBar1.Panels(1).Text = "Seleccionando servicios"
    ElseIf CodigoGrupo = 0 Then
        StatusBar1.Panels(1).Text = "Seleccione un grupo"
    ElseIf ConsultaPendiente Then
        StatusBar1.Panels(1).Text = "Consulta pendiente"
    ElseIf Not ConsultaValida Then
        StatusBar1.Panels(1).Text = "Pulse Cargar Datos del Grupo"
    ElseIf Len(txtBuscar.Text) > 0 Then
        StatusBar1.Panels(1).Text = "Consulta filtrada"
    Else
        StatusBar1.Panels(1).Text = "Consulta de documentos"
    End If
    If CodigoGrupo > 0 Then
        StatusBar1.Panels(2).Text = "Grupo " & CStr(CodigoGrupo) & " - " & LblNom.Caption
    Else
        StatusBar1.Panels(2).Text = "Buscar o crear un grupo"
    End If
    If Seleccionando Then
        StatusBar1.Panels(3).Text = "Marcados: " & CStr(Marcados) & " de " & CStr(ViewAgrupar.Nodes.Count)
    ElseIf ConsultaPendiente Then
        StatusBar1.Panels(3).Text = "Recargue los datos del grupo"
    ElseIf ConsultaValida Then
        StatusBar1.Panels(3).Text = "Servicios: " & CStr(GridAgrupar.Rows - 1) & " | Documentos: " & CStr(Grid.Rows - 1) & " de " & CStr(RsRegist.RecordCount)
    Else
        StatusBar1.Panels(3).Text = "Servicios: " & CStr(GridAgrupar.Rows - 1) & " | Documentos sin consultar"
    End If
End Sub

Private Sub LimpiarDocumentos()
    ConsultaValida = False
    If Not RsRegist Is Nothing Then
        If RsRegist.State <> adStateClosed Then RsRegist.Close
        Set RsRegist = Nothing
    End If
    Grid.Rows = 1: GridSta.Rows = 1: lblCount.Caption = "0"
End Sub

Private Sub TerminarSeleccion()
    Seleccionando = False
    ViewAgrupar.Visible = False: GridAgrupar.Visible = True
    ViewAgrupar.Nodes.Clear
    ActualizarEstado
End Sub

Public Sub CargarGrupo(ByVal Codigo As Long)
    Dim Datos As ADODB.Recordset, Servicios As ADODB.Recordset
    Dim Nombre As String, Columna As Integer, Fila As Long
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Codigo <= 0 Then Err.Raise 5, , "Seleccione un grupo valido."
    'Consultar antes de sustituir el grupo que se muestra en pantalla.
    Set Datos = ConsultarDatos("SELECT Item,NomGrp FROM dbo.MdbGruposServ WHERE Item=?", CStr(Codigo), True)
    If Datos.EOF Then Err.Raise 5, , "El grupo no existe."
    Nombre = CStr(Datos.Fields(1).Value): Datos.Close
    Set Servicios = ConsultarDatos("SELECT CodSer,NroBol,NomCli FROM dbo.ViewDatosServicios WHERE Grupo=? ORDER BY CodSer DESC", CStr(Codigo), True)
    Cargando = True
    LimpiarDocumentos
    CodigoGrupo = Codigo: LblBol.Caption = CStr(Codigo): LblNom.Caption = Nombre
    GridAgrupar.Rows = 1
    Do Until Servicios.EOF
        GridAgrupar.Rows = GridAgrupar.Rows + 1: Fila = GridAgrupar.Rows - 1
        For Columna = 0 To 2: GridAgrupar.TextMatrix(Fila, Columna) = TextoDato(Servicios.Fields(Columna).Value): Next
        Servicios.MoveNext
    Loop
    Servicios.Close
    txtBuscar.Text = "": CmbBuscar.ListIndex = 1
    ConsultaPendiente = False: Cargando = False
    TerminarSeleccion
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    Cargando = False
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    If Not Servicios Is Nothing Then Servicios.Close
    On Error GoTo 0
    Err.Raise Numero, "CargarGrupo", Descripcion
End Sub

Private Sub ActualizarGrupoGuardado(ByVal Codigo As Long, ByVal Detalle As String)
    Dim Descripcion As String
    On Error GoTo ErrInt
    CargarGrupo Codigo
    CargarDocumentos
    Exit Sub
ErrInt:
    Descripcion = Err.Description
    If CodigoGrupo <> Codigo Then LblNom.Caption = ""
    CodigoGrupo = Codigo: LblBol.Caption = CStr(Codigo)
    ConsultaPendiente = True
    LimpiarDocumentos: GridAgrupar.Rows = 1
    TerminarSeleccion
    MsgBox Detalle & vbCrLf & "La operacion ya esta guardada, pero no se pudo actualizar la consulta: " & Descripcion & vbCrLf & "No repita el guardado. Use Cargar Datos del Grupo para recuperar la consulta.", vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdBuscar_Click()
    On Error GoTo ErrInt
    If Not PuedeCerrar() Then Exit Sub
    TerminarSeleccion
    Var.Varflotante = 3
    Load FrmMenuFlot
    FrmMenuFlot.Move MDIPri.Left + (MDIPri.Width - FrmMenuFlot.Width) / 2, MDIPri.Top + (MDIPri.Height - FrmMenuFlot.Height) / 2
    FrmMenuFlot.Show vbModal, MDIPri
    If Var.RegistroSeleccionado Then CargarGrupo CLng(Var.ResultadoRegistro(0))
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdNewGrup_Click()
    Dim Codigo As Long
    On Error GoTo ErrInt
    If Not PuedeCerrar() Then Exit Sub
    TerminarSeleccion
    Load FrmNewGrup
    FrmNewGrup.CodigoCreado = 0
    FrmNewGrup.Show vbModal, MDIPri
    Codigo = FrmNewGrup.CodigoCreado
    Unload FrmNewGrup
    If Codigo > 0 Then ActualizarGrupoGuardado Codigo, "Grupo creado: " & CStr(Codigo)
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdAgrupar_Click()
    Dim Datos As ADODB.Recordset, Nodo As MSComctlLib.Node
    On Error GoTo ErrInt
    If CodigoGrupo = 0 Or Guardando Or Seleccionando Or ConsultaPendiente Then Exit Sub
    ViewAgrupar.Nodes.Clear
    Set Datos = ConsultarDatos("SELECT CodSer,NroBol,NomCli FROM dbo.ViewDatosServicios WHERE Grupo=0 ORDER BY CodSer DESC")
    Do Until Datos.EOF
        Set Nodo = ViewAgrupar.Nodes.Add(, , "S" & CStr(Datos.Fields(0).Value), Format$(Datos.Fields(0).Value, "000000") & "-" & Format$(Datos.Fields(1).Value, "000000") & "-" & CStr(Datos.Fields(2).Value))
        Nodo.Tag = CStr(Datos.Fields(0).Value)
        Datos.MoveNext
    Loop
    Datos.Close
    Seleccionando = True
    ViewAgrupar.Visible = True: GridAgrupar.Visible = False
    ActualizarEstado
    If ViewAgrupar.Nodes.Count = 0 Then MsgBox "No hay servicios sin grupo disponibles para agrupar.", vbInformation, "Clever SoftPro"
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub ViewAgrupar_NodeCheck(ByVal Node As MSComctlLib.Node)
    ActualizarEstado
End Sub

Private Sub CmdAgru_Click()
    Dim Codigos As New Collection, Nodo As MSComctlLib.Node
    On Error GoTo ErrInt
    If Not Seleccionando Or Guardando Or CodigoGrupo = 0 Then Exit Sub
    For Each Nodo In ViewAgrupar.Nodes
        If Nodo.Checked Then Codigos.Add CStr(Nodo.Tag)
    Next
    If Codigos.Count = 0 Then Exit Sub
    Guardando = True: ActualizarEstado
    AgruparServicios CodigoGrupo, 1, Codigos
    ActualizarGrupoGuardado CodigoGrupo, CStr(Codigos.Count) & " servicios agrupados."
    Guardando = False: ActualizarEstado
    Exit Sub
ErrInt:
    Guardando = False: ActualizarEstado
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdCanAgru_Click()
    If Guardando Then Exit Sub
    If PuedeCerrar() Then TerminarSeleccion
End Sub

Private Sub GridAgrupar_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim Codigos As New Collection, Codigo As String
    On Error GoTo ErrInt
    If KeyCode <> vbKeyDelete Then Exit Sub
    KeyCode = 0
    If CodigoGrupo = 0 Or Guardando Or Seleccionando Or ConsultaPendiente Then Exit Sub
    If GridAgrupar.Row < 1 Or GridAgrupar.Rows < 2 Then Exit Sub
    Codigo = GridAgrupar.TextMatrix(GridAgrupar.Row, 0)
    If Not CodigoValido(Codigo) Then Exit Sub
    If MsgBox("Se retirara el servicio del siguiente grupo:" & vbCrLf & vbCrLf & _
        "Grupo: " & CStr(CodigoGrupo) & " - " & LblNom.Caption & vbCrLf & _
        "Servicio: " & Codigo & vbCrLf & _
        "Boleta: " & GridAgrupar.TextMatrix(GridAgrupar.Row, 1) & vbCrLf & _
        "Cliente: " & GridAgrupar.TextMatrix(GridAgrupar.Row, 2) & vbCrLf & vbCrLf & _
        "Se conservaran el servicio, sus documentos y su historial." & vbCrLf & _
        "Desea continuar?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Retirar Servicio del Grupo") <> vbYes Then Exit Sub
    Codigos.Add Codigo
    Guardando = True: ActualizarEstado
    AgruparServicios CodigoGrupo, 2, Codigos
    ActualizarGrupoGuardado CodigoGrupo, "Servicio " & Codigo & " retirado del grupo."
    Guardando = False: ActualizarEstado
    Exit Sub
ErrInt:
    Guardando = False: ActualizarEstado
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdMostrar_Click()
    On Error GoTo ErrInt
    If CodigoGrupo = 0 Or Guardando Or Seleccionando Then Exit Sub
    If ConsultaPendiente Then CargarGrupo CodigoGrupo
    CargarDocumentos
    Exit Sub
ErrInt:
    ConsultaPendiente = True
    LimpiarDocumentos
    ActualizarEstado
    MsgBox "No se pudo cargar la consulta: " & Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CargarDocumentos(Optional ByVal Consultar As Boolean = True)
    Dim Datos As ADODB.Recordset, Fila As Long, Columna As Integer, Posicion As Long
    Dim Estado As String, Estadisticas As Object, Clave As Variant, Mostrar As Boolean, Campo As String
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    If Consultar Then
        LimpiarDocumentos
        Set RsRegist = ConsultarDatos("SELECT Codigo,Producto,Empresa,Direccion,Distrito,Contacto,Descripcion,Servicio,Boleta,Cliente,UltimoSta,Fecha,Informe FROM dbo.ViewGrupos WHERE Grupo=? ORDER BY Codigo", CStr(CodigoGrupo), True)
    End If
    If RsRegist Is Nothing Then Exit Sub
    Set Datos = RsRegist.Clone
    If Not Datos.EOF Then Datos.MoveFirst
    Set Estadisticas = CreateObject("Scripting.Dictionary")
    Estadisticas.CompareMode = vbTextCompare
    Grid.Rows = 1: GridSta.Rows = 1
    Campo = CmbBuscar.Text
    Do Until Datos.EOF
        Mostrar = True
        If Campo <> "" And txtBuscar.Text <> "" Then Mostrar = (InStr(1, TextoDato(Datos.Fields(Campo).Value), txtBuscar.Text, vbTextCompare) > 0)
        If Mostrar Then
            Grid.Rows = Grid.Rows + 1: Fila = Grid.Rows - 1
            For Columna = 0 To 12: Grid.TextMatrix(Fila, Columna) = TextoDato(Datos.Fields(Columna).Value): Next
            Estado = Grid.TextMatrix(Fila, 10)
            If Len(Estado) = 0 Then Estado = "Pendiente"
            If Not Estadisticas.Exists(Estado) Then Estadisticas.Add Estado, 0
            Estadisticas(Estado) = CLng(Estadisticas(Estado)) + 1
        End If
        Datos.MoveNext
    Loop
    Datos.Close
    For Each Clave In Estadisticas.Keys
        GridSta.Rows = GridSta.Rows + 1: Posicion = GridSta.Rows - 1
        GridSta.TextMatrix(Posicion, 0) = CStr(Clave): GridSta.TextMatrix(Posicion, 1) = Format$(Estadisticas(Clave), "00")
    Next
    lblCount.Caption = CStr(Grid.Rows - 1)
    ConsultaValida = True: ConsultaPendiente = False
    ActualizarEstado
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    LimpiarDocumentos
    On Error GoTo 0
    ConsultaPendiente = True
    ActualizarEstado
    Err.Raise Numero, "CargarDocumentos", Descripcion
End Sub

Private Sub txtBuscar_Change()
    On Error GoTo ErrInt
    If Not Cargando And CodigoGrupo > 0 And ConsultaValida And Not Seleccionando And Not Guardando Then CargarDocumentos False
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmbBuscar_Click()
    txtBuscar_Change
End Sub

Private Sub CmdGenExcel_Click()
    Dim Contexto As String
    On Error GoTo ErrInt
    If Not ConsultaValida Or ConsultaPendiente Or Seleccionando Or Guardando Or CodigoGrupo = 0 Or Grid.Rows < 2 Then Exit Sub
    Contexto = "Grupo: " & CStr(CodigoGrupo) & " - " & LblNom.Caption
    If Len(txtBuscar.Text) > 0 Then Contexto = Contexto & " | Filtro: " & CmbBuscar.Text & " contiene " & txtBuscar.Text Else Contexto = Contexto & " | Sin filtro"
    MdiExport.CmdGenExcel Grid, "REPORTE DE SERVICIOS", , Contexto
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not ConsultaValida Or Seleccionando Or Guardando Then Exit Sub
    If Grid.MouseRow <> 0 Or Grid.Rows < 3 Or Grid.MouseCol < 0 Then Exit Sub
    Grid.Col = Grid.MouseCol: Grid.ColSel = Grid.Col: Grid.Row = 1: Grid.RowSel = Grid.Rows - 1
    Grid.Sort = Orden
    If Orden = 1 Then Orden = 2 Else Orden = 1
End Sub

Public Function PuedeCerrar() As Boolean
    If Guardando Then Exit Function
    PuedeCerrar = True
    If Not Seleccionando Then Exit Function
    If ServiciosMarcados() > 0 Then
        PuedeCerrar = (MsgBox("Hay servicios seleccionados sin agrupar. Desea descartar la seleccion?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") = vbYes)
        If PuedeCerrar Then TerminarSeleccion
    End If
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Seleccionando Then KeyCode = 0: CmdCanAgru_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not CerrandoAplicacion Then Cancel = Not PuedeCerrar()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    LimpiarDocumentos
End Sub

Private Sub Form_Activate()
    VerificarAccesoFormulario Me, 4
End Sub
