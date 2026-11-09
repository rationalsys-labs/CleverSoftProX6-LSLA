VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuracion de Informaciones"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10200
   KeyPreview = -1
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   10200
   Begin VB.Frame FraSta 
      Caption         =   "Configuracion de Status"
      Height          =   3795
      Left            =   3825
      TabIndex        =   2
      Top             =   45
      Visible         =   0   'False
      Width           =   6225
      Begin VB.CommandButton CmdCancelar
         Caption = "Cancelar"
         Left = 4905
         Top = 2700
         Width = 1140
         Height = 330
         TabIndex = 11
      End
      Begin VB.OptionButton OptModSta 
         Caption         =   "Modificar Status"
         Height          =   195
         Left            =   3780
         TabIndex        =   10
         Top             =   1395
         Width           =   1455
      End
      Begin VB.OptionButton OptNewSta 
         Caption         =   "Nuevo Status"
         Height          =   195
         Left            =   3780
         TabIndex        =   9
         Top             =   315
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.CommandButton CmdModSta 
         Caption         =   "Modificar"
         Enabled         =   0   'False
         Height          =   330
         Left            =   4905
         TabIndex        =   8
         Top             =   1980
         Width           =   1140
      End
      Begin VB.TextBox TxtModSta 
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         Height          =   285
         Left            =   3825
         TabIndex        =   7
         Top             =   1620
         Width           =   2220
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Dar de baja"
         Height          =   330
         Left            =   3825
         TabIndex        =   6
         Top             =   3285
         Width           =   1185
      End
      Begin VB.CommandButton CmdNewSta 
         Caption         =   "Crear"
         Height          =   330
         Left            =   4905
         TabIndex        =   5
         Top             =   900
         Width           =   1140
      End
      Begin VB.TextBox TxtNewSta 
         Height          =   285
         Left            =   3825
         TabIndex        =   4
         Top             =   540
         Width           =   2220
      End
      Begin MSFlexGridLib.MSFlexGrid GridSta 
         Height          =   3300
         Left            =   225
         TabIndex        =   3
         Top             =   315
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   5821
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
      End
   End
   Begin MSComctlLib.TreeView ViewMenus 
      Height          =   3690
      Left            =   135
      TabIndex        =   0
      Top             =   135
      Width           =   3570
      _ExtentX        =   6297
      _ExtentY        =   6509
      _Version        =   393217
      Indentation     =   90
      LineStyle       =   1
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   3930
      Width           =   10200
      _ExtentX        =   17992
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
End
Attribute VB_Name = "FrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CodigoEdicion As Long
Private NombreAnterior As String
Private ConsultandoDistritos As Boolean
Public CodigoGuardado As Long

Private CargaValida As Boolean
Private Cargando As Boolean

Private Sub Form_Load()
    On Error GoTo ErrInt
    Me.Icon = MDIPri.ICO.ListImages(25).Picture
    ViewMenus.ImageList = MDIPri.ICO
    ViewMenus.Nodes.Add , , "Acciones", "Acciones", MDIPri.ICO.ListImages(23).Key
    ViewMenus.Nodes.Add "Acciones", tvwChild, "raiz1", "Configurar Status de Registros", MDIPri.ICO.ListImages(24).Key
    ViewMenus.Nodes.Add "Acciones", tvwChild, "raiz2", "Configurar Distritos de Sistema", MDIPri.ICO.ListImages(24).Key
    ViewMenus.Nodes(1).Expanded = True
    TxtNewSta.MaxLength = 20: TxtModSta.MaxLength = 20
    Dim Panel As MSComctlLib.Panel
    For Each Panel In StatusBar1.Panels: Panel.MinWidth = 15: Next
    StatusBar1.Panels(1).Width = 2400: StatusBar1.Panels(2).AutoSize = sbrSpring: StatusBar1.Panels(3).Width = 2100
    CargarEstados
    FraSta.Visible = True
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Public Sub CargarEstados()
    Dim Datos As ADODB.Recordset, Fila As Long
    On Error GoTo ErrorConsulta
    CargaValida = False: Cargando = True
    ConsultandoDistritos = False: CodigoEdicion = 0: NombreAnterior = ""
    GridSta.Rows = 1: GridSta.Cols = 2
    GridSta.TextMatrix(0, 0) = "Item": GridSta.ColWidth(0) = 600
    GridSta.TextMatrix(0, 1) = "Status": GridSta.ColWidth(1) = 2500
    Set Datos = ConsultarDatos("SELECT ITEM,Status FROM dbo.IntStatus WHERE ActEst=1 AND Active=1 ORDER BY ITEM")
    Do Until Datos.EOF
        GridSta.Rows = GridSta.Rows + 1: Fila = GridSta.Rows - 1
        GridSta.TextMatrix(Fila, 0) = CStr(Datos.Fields(0).Value): GridSta.TextMatrix(Fila, 1) = CStr(Datos.Fields(1).Value)
        Datos.MoveNext
    Loop
    Datos.Close
    TxtNewSta.Text = "": TxtModSta.Text = ""
    CargaValida = True: Cargando = False
    MostrarEdicion True
    Exit Sub
ErrorConsulta:
    Dim Numero As Long, Descripcion As String
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    GridSta.Rows = 1: CodigoEdicion = 0: Cargando = False: CargaValida = False
    ActualizarModo
    On Error GoTo 0
    Err.Raise Numero, , Descripcion
End Sub

Public Sub ConsultarDistritos()
    Dim Datos As ADODB.Recordset, Fila As Long, Columna As Integer
    On Error GoTo ErrorConsulta
    CargaValida = False: Cargando = True
    ConsultandoDistritos = True: CodigoEdicion = 0
    GridSta.Rows = 1: GridSta.Cols = 3
    GridSta.TextMatrix(0, 0) = "Codigo": GridSta.ColWidth(0) = 600
    GridSta.TextMatrix(0, 1) = "Postal": GridSta.ColWidth(1) = 650
    GridSta.TextMatrix(0, 2) = "Distrito": GridSta.ColWidth(2) = 1850
    Set Datos = ConsultarDatos("SELECT CoDist,CodPos,NomDis FROM dbo.MDBDistritos WHERE Active=1 ORDER BY NomDis")
    Do Until Datos.EOF
        GridSta.Rows = GridSta.Rows + 1: Fila = GridSta.Rows - 1
        For Columna = 0 To 2: GridSta.TextMatrix(Fila, Columna) = TextoDato(Datos.Fields(Columna).Value): Next
        Datos.MoveNext
    Loop
    Datos.Close
    CargaValida = True: Cargando = False
    MostrarEdicion False
    Exit Sub
ErrorConsulta:
    Dim Numero As Long, Descripcion As String
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    GridSta.Rows = 1: CodigoEdicion = 0: Cargando = False: CargaValida = False
    ActualizarModo
    On Error GoTo 0
    Err.Raise Numero, , Descripcion
End Sub

Private Sub MostrarEdicion(ByVal Mostrar As Boolean)
    FraSta.Caption = IIf(Mostrar, "Configuracion de Status", "Distritos de Sistema")
    OptNewSta.Visible = Mostrar: OptModSta.Visible = Mostrar: TxtNewSta.Visible = Mostrar: TxtModSta.Visible = Mostrar
    CmdNewSta.Visible = Mostrar: CmdModSta.Visible = Mostrar: Command2.Visible = Mostrar
    ActualizarModo
End Sub

Private Sub ActualizarModo()
    If Cargando Then Exit Sub
    TxtNewSta.Enabled = (CargaValida And Not ConsultandoDistritos And OptNewSta.Value)
    CmdNewSta.Enabled = (TxtNewSta.Enabled And Len(Trim$(TxtNewSta.Text)) > 0)
    TxtModSta.Enabled = (CargaValida And Not ConsultandoDistritos And OptModSta.Value And CodigoEdicion > 0)
    CmdModSta.Enabled = (TxtModSta.Enabled And Len(Trim$(TxtModSta.Text)) > 0 And TxtModSta.Text <> NombreAnterior)
    Command2.Enabled = (CargaValida And Not ConsultandoDistritos And CodigoEdicion > 2)
    CmdCancelar.Visible = Not ConsultandoDistritos
    If TxtNewSta.Enabled Then TxtNewSta.BackColor = vbWindowBackground Else TxtNewSta.BackColor = &HC0C0C0
    If TxtModSta.Enabled Then TxtModSta.BackColor = vbWindowBackground Else TxtModSta.BackColor = &HC0C0C0
    If Not CargaValida Then
        StatusBar1.Panels(1).Text = "Consulta no disponible"
    ElseIf ConsultandoDistritos Then
        StatusBar1.Panels(1).Text = "Distritos: solo consulta"
    Else
        StatusBar1.Panels(1).Text = "Configuracion de estados"
    End If
    If CodigoEdicion > 0 Then StatusBar1.Panels(2).Text = CStr(CodigoEdicion) & " - " & NombreAnterior Else StatusBar1.Panels(2).Text = "Seleccione un estado para modificar o dar de baja"
    StatusBar1.Panels(3).Text = CStr(GridSta.Rows - 1) & " registros"
End Sub

Public Sub SeleccionarEstado(ByVal Codigo As Long)
    Dim Fila As Long
    If Not PuedeCerrar() Then Exit Sub
    For Fila = 1 To GridSta.Rows - 1
        If CLng(GridSta.TextMatrix(Fila, 0)) = Codigo Then
            GridSta.Row = Fila: CodigoEdicion = Codigo
            NombreAnterior = GridSta.TextMatrix(Fila, 1): TxtModSta.Text = NombreAnterior
            ActualizarModo
            Exit Sub
        End If
    Next
    Err.Raise 5, , "Seleccione un estado de la lista."
End Sub

Public Sub CrearEstado()
    CodigoGuardado = MdiCatalogos.GuardarEstado(0, 1, TxtNewSta.Text)
    CargarEstados
End Sub

Public Sub ModificarEstado()
    If CodigoEdicion = 0 Or ConsultandoDistritos Then Err.Raise 5, , "Seleccione el estado que desea modificar."
    CodigoGuardado = MdiCatalogos.GuardarEstado(CodigoEdicion, 2, TxtModSta.Text)
    CargarEstados
End Sub

Public Sub DarBajaEstado()
    If CodigoEdicion = 0 Or ConsultandoDistritos Then Err.Raise 5, , "Seleccione el estado que desea dar de baja."
    CodigoGuardado = MdiCatalogos.GuardarEstado(CodigoEdicion, 3)
    CargarEstados
End Sub

Private Sub CmdNewSta_Click()
    On Error GoTo ErrInt
    CrearEstado
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdModSta_Click()
    On Error GoTo ErrInt
    ModificarEstado
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub Command2_Click()
    On Error GoTo ErrInt
    If CodigoEdicion <= 2 Or ConsultandoDistritos Then Err.Raise 5, , "Seleccione un estado distinto de los basicos 1 y 2."
    If MsgBox("Desea dar de baja el estado " & CStr(CodigoEdicion) & " - " & NombreAnterior & "? Se conservara su historial.", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") <> vbYes Then Exit Sub
    DarBajaEstado
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub GridSta_Click()
    If Not ConsultandoDistritos And GridSta.Row > 0 And GridSta.Rows > 1 Then SeleccionarEstado CLng(GridSta.TextMatrix(GridSta.Row, 0))
End Sub

Private Sub OptNewSta_Click()
    ActualizarModo
End Sub

Private Sub OptModSta_Click()
    ActualizarModo
End Sub

Private Sub ViewMenus_Click()
    On Error GoTo ErrInt
    If ViewMenus.SelectedItem Is Nothing Then Exit Sub
    If Not PuedeCerrar() Then Exit Sub
    Select Case ViewMenus.SelectedItem.Key
        Case "raiz1": CargarEstados: FraSta.Visible = True
        Case "raiz2": ConsultarDistritos: FraSta.Visible = True
        Case Else: FraSta.Visible = False
    End Select
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Public Function PuedeCerrar() As Boolean
    PuedeCerrar = True
    If ConsultandoDistritos Then Exit Function
    If Trim$(TxtNewSta.Text) = "" And TxtModSta.Text = NombreAnterior Then Exit Function
    PuedeCerrar = (MsgBox("Hay cambios de estado sin guardar. Desea descartarlos?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") = vbYes)
    If PuedeCerrar Then TxtNewSta.Text = "": TxtModSta.Text = NombreAnterior
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not CerrandoAplicacion Then
        If Not PuedeCerrar() Then Cancel = True
    End If
End Sub
Private Sub TxtNewSta_Change()
    ActualizarModo
End Sub
Private Sub TxtModSta_Change()
    ActualizarModo
End Sub
Private Sub CmdCancelar_Click()
    If PuedeCerrar() Then CargarEstados
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then KeyCode = 0: CmdCancelar_Click
End Sub
Private Sub Form_Activate()
    VerificarAccesoFormulario Me, 10
End Sub
