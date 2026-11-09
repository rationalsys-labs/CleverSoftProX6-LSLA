VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Inform_Status 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Informe de Status"
   ClientHeight    =   5865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5865
   ScaleWidth      =   12255
   Begin VB.CommandButton CmdgenExcel 
      Caption         =   "Generar Excel"
      Height          =   420
      Left            =   10395
      TabIndex        =   25
      Top             =   5130
      Width           =   1770
   End
   Begin VB.Frame Frame1 
      Caption         =   "Filtros"
      Height          =   2760
      Left            =   5445
      TabIndex        =   6
      Top             =   495
      Width           =   6765
      Begin VB.ComboBox CmbCol 
         Height          =   315
         Index           =   0
         Left            =   225
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   270
         Width           =   1680
      End
      Begin VB.Frame Frame5 
         Height          =   375
         Left            =   90
         TabIndex        =   22
         Top             =   2205
         Width           =   6405
         Begin VB.ComboBox CmbCol 
            Height          =   315
            Index           =   3
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   0
            Width           =   1680
         End
         Begin VB.TextBox TxtFil 
            BackColor       =   &H80000003&
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   3645
            TabIndex        =   24
            Top             =   0
            Width           =   2625
         End
         Begin VB.ComboBox CmbFil 
            BackColor       =   &H80000003&
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   0
            Width           =   1680
         End
      End
      Begin VB.Frame Frame4 
         Height          =   645
         Left            =   90
         TabIndex        =   17
         Top             =   1530
         Width           =   6405
         Begin VB.ComboBox CmbCol 
            Height          =   315
            Index           =   2
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   0
            Width           =   1680
         End
         Begin VB.OptionButton OptO 
            Caption         =   "O"
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   1035
            TabIndex        =   21
            Top             =   360
            Width           =   510
         End
         Begin VB.OptionButton OptY 
            Caption         =   "Y"
            Enabled         =   0   'False
            Height          =   240
            Index           =   2
            Left            =   450
            TabIndex        =   20
            Top             =   360
            Value           =   -1  'True
            Width           =   510
         End
         Begin VB.TextBox TxtFil 
            BackColor       =   &H80000003&
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   3645
            TabIndex        =   19
            Top             =   0
            Width           =   2625
         End
         Begin VB.ComboBox CmbFil 
            BackColor       =   &H80000003&
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   18
            Top             =   0
            Width           =   1680
         End
      End
      Begin VB.Frame Frame3 
         Height          =   645
         Left            =   90
         TabIndex        =   12
         Top             =   900
         Width           =   6405
         Begin VB.ComboBox CmbCol 
            Height          =   315
            Index           =   1
            Left            =   135
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   0
            Width           =   1680
         End
         Begin VB.OptionButton OptO 
            Caption         =   "O"
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   1035
            TabIndex        =   16
            Top             =   360
            Width           =   510
         End
         Begin VB.OptionButton OptY 
            Caption         =   "Y"
            Enabled         =   0   'False
            Height          =   240
            Index           =   1
            Left            =   450
            TabIndex        =   15
            Top             =   360
            Value           =   -1  'True
            Width           =   510
         End
         Begin VB.TextBox TxtFil 
            BackColor       =   &H80000003&
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   3645
            TabIndex        =   14
            Top             =   0
            Width           =   2625
         End
         Begin VB.ComboBox CmbFil 
            BackColor       =   &H80000003&
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   0
            Width           =   1680
         End
      End
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   90
         TabIndex        =   7
         Top             =   270
         Width           =   6405
         Begin VB.OptionButton OptO 
            Caption         =   "O"
            Height          =   240
            Index           =   0
            Left            =   1035
            TabIndex        =   11
            Top             =   360
            Width           =   510
         End
         Begin VB.OptionButton OptY 
            Caption         =   "Y"
            Height          =   240
            Index           =   0
            Left            =   450
            TabIndex        =   10
            Top             =   360
            Value           =   -1  'True
            Width           =   510
         End
         Begin VB.TextBox TxtFil 
            Height          =   315
            Index           =   0
            Left            =   3645
            TabIndex        =   9
            Top             =   0
            Width           =   2625
         End
         Begin VB.ComboBox CmbFil 
            Height          =   315
            Index           =   0
            Left            =   1890
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   0
            Width           =   1680
         End
      End
   End
   Begin VB.ComboBox CmbOrigen 
      Height          =   315
      Left            =   1530
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   90
      Width           =   3795
   End
   Begin VB.CommandButton CmdImprimir 
      Caption         =   "Imprimir Informe"
      Height          =   420
      Left            =   10395
      TabIndex        =   3
      Top             =   4680
      Width           =   1770
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Guardar Informe"
      Height          =   420
      Left            =   5445
      TabIndex        =   2
      Top             =   5130
      Width           =   1770
   End
   Begin MSComctlLib.TreeView ViewMenus 
      Height          =   4950
      Left            =   90
      TabIndex        =   0
      Top             =   585
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   8731
      _Version        =   393217
      Indentation     =   90
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   1
      Top             =   5625
      Width           =   12255
      _ExtentX        =   21616
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
   Begin VB.Label Label1 
      Caption         =   "Origen de Reporte"
      Height          =   240
      Left            =   90
      TabIndex        =   5
      Top             =   135
      Width           =   1365
   End
   Begin MSComDlg.CommonDialog CUAD
      Left = 0
      Top = 0
      _ExtentX = 847
      _ExtentY = 847
      _Version = 393216
   End
End
Attribute VB_Name = "Inform_Status"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Preparando As Boolean

Private Sub Form_Load()
    Dim Posicion As Integer, Operadores As Variant, Operador As Variant
    Me.Icon = MDIPri.ICO.ListImages(25).Picture
    ViewMenus.ImageList = MDIPri.ICO
    CmbOrigen.AddItem "Registros de Status": CmbOrigen.AddItem "Registros de Servicios"
    CmbOrigen.AddItem "registros de Clientes": CmbOrigen.AddItem "registro de Documentos"
    Operadores = Array("Es igual a", "No es igual a", "Es mayor que", "Es mayor o igual a", "Es menor que", "Es menor o igual a", "Comienza por", "No comienza por", "Termina Con", "No termina con", "Contiene", "No Contiene")
    For Posicion = 0 To 3
        For Each Operador In Operadores: CmbFil(Posicion).AddItem CStr(Operador): Next
    Next
    Dim Panel As MSComctlLib.Panel
    For Each Panel In StatusBar1.Panels: Panel.MinWidth = 15: Next
    StatusBar1.Panels(1).Width = 2700: StatusBar1.Panels(2).AutoSize = sbrSpring: StatusBar1.Panels(3).Width = 2400
    HabilitarFiltros
End Sub

Private Sub CmbOrigen_Click()
    Dim Datos As ADODB.Recordset, Campo As ADODB.Field, Posicion As Integer, Etiquetas As Variant, Indice As Integer
    On Error GoTo ErrInt
    Preparando = True: ViewMenus.Nodes.Clear
    For Posicion = 0 To 3
        CmbCol(Posicion).Clear: CmbFil(Posicion).ListIndex = -1: TxtFil(Posicion).Text = ""
    Next
    Set Datos = ConsultarDatos("SELECT TOP(0) * FROM (" & ConsultaOrigen(CmbOrigen.ListIndex + 1) & ") Datos")
    Etiquetas = Array("Codigo de Registro", "Emision del Registro", "Producto del Registro", "Documento Registro del Cliente", "Nro de Servicio", "Nro de Boleta del Servicio", "Producto del Servicio", "Empresa que Requirio el Servicio", "Cantidad de Unidades del servicio", "Destinos del Servicio", "Nombre del Destinatario", "Direccion del Destinatario", "Distrito del Destinatario", "Contacto del Destinatario", "Status del Registro", "Fecha del Registro", "Informe de Status")
    For Each Campo In Datos.Fields
        If CmbOrigen.ListIndex = 0 Then
            ViewMenus.Nodes.Add , , Campo.Name, Etiquetas(Indice), MDIPri.ICO.ListImages(31).Key
        Else
            ViewMenus.Nodes.Add , , Campo.Name, Campo.Name, MDIPri.ICO.ListImages(31).Key
        End If
        Indice = Indice + 1
    Next
    Datos.Close
    Preparando = False: HabilitarFiltros
    Exit Sub
ErrInt:
    ViewMenus.Nodes.Clear
    Preparando = False: HabilitarFiltros
    Dim Descripcion As String
    Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    On Error GoTo 0
    MsgBox Descripcion, vbExclamation, "Clever SoftPro"
End Sub

Public Sub ActualizarColumnas()
    Dim Posicion As Integer, Indice As Integer, Anterior As String, Nodo As MSComctlLib.Node
    Preparando = True
    For Posicion = 0 To 3
        Anterior = CmbCol(Posicion).Text: CmbCol(Posicion).Clear
        For Each Nodo In ViewMenus.Nodes
            If Nodo.Checked Then CmbCol(Posicion).AddItem Nodo.Key
        Next
        For Indice = 0 To CmbCol(Posicion).ListCount - 1
            If CmbCol(Posicion).List(Indice) = Anterior Then CmbCol(Posicion).ListIndex = Indice: Exit For
        Next
        If CmbCol(Posicion).ListIndex < 0 Then CmbFil(Posicion).ListIndex = -1: TxtFil(Posicion).Text = ""
    Next
    Preparando = False: HabilitarFiltros
End Sub

Private Sub HabilitarFiltros()
    Dim Posicion As Integer, Habilitar As Boolean
    Preparando = True
    Habilitar = (CmbOrigen.ListIndex >= 0 And CmbCol(0).ListCount > 0)
    For Posicion = 0 To 3
        CmbCol(Posicion).Enabled = Habilitar: CmbFil(Posicion).Enabled = Habilitar: TxtFil(Posicion).Enabled = Habilitar
        If Habilitar Then CmbFil(Posicion).BackColor = vbWindowBackground: TxtFil(Posicion).BackColor = vbWindowBackground Else CmbFil(Posicion).BackColor = &H80000003: TxtFil(Posicion).BackColor = &H80000003
        If Posicion < 3 Then OptY(Posicion).Enabled = Habilitar: OptO(Posicion).Enabled = Habilitar
        Habilitar = Habilitar And CmbCol(Posicion).ListIndex >= 0 And CmbFil(Posicion).ListIndex >= 0
        If Not Habilitar And Posicion < 3 Then
            CmbCol(Posicion + 1).ListIndex = -1: CmbFil(Posicion + 1).ListIndex = -1: TxtFil(Posicion + 1).Text = ""
        End If
    Next
    Preparando = False
    ActualizarAcciones
End Sub

Private Sub ViewMenus_NodeCheck(ByVal Node As MSComctlLib.Node)
    If Not Preparando Then ActualizarColumnas
End Sub

Private Sub CmbCol_Click(Index As Integer)
    If Not Preparando Then HabilitarFiltros
End Sub

Private Sub CmbFil_Click(Index As Integer)
    If Not Preparando Then HabilitarFiltros
End Sub

Public Function GenerarInforme() As ADODB.Recordset
    Dim Columnas As New Collection, Filtros(0 To 3) As FiltroInforme, Nodo As MSComctlLib.Node, Posicion As Integer, Cantidad As Integer
    For Each Nodo In ViewMenus.Nodes
        If Nodo.Checked Then Columnas.Add Nodo.Key
    Next
    For Posicion = 0 To 3
        If CmbCol(Posicion).Enabled Then
            If CmbCol(Posicion).ListIndex >= 0 Or CmbFil(Posicion).ListIndex >= 0 Or TxtFil(Posicion).Text <> "" Then
                If CmbCol(Posicion).ListIndex < 0 Or CmbFil(Posicion).ListIndex < 0 Then Err.Raise 5, , "Complete el campo y la condicion de cada filtro."
                Filtros(Cantidad).Campo = CmbCol(Posicion).Text
                Filtros(Cantidad).Operador = CmbFil(Posicion).ListIndex + 1: Filtros(Cantidad).Valor = TxtFil(Posicion).Text
                If Posicion < 3 Then Filtros(Cantidad).OpcionO = OptO(Posicion).Value
                Cantidad = Cantidad + 1
            End If
        End If
    Next
    Set GenerarInforme = ConsultarInforme(CmbOrigen.ListIndex + 1, Columnas, Filtros, Cantidad)
End Function

Public Sub ExportarInforme(Optional ByVal Ruta As String = "")
    Dim Datos As ADODB.Recordset, Numero As Long, Descripcion As String
    On Error GoTo ErrExportar
    Set Datos = GenerarInforme()
    StatusBar1.Panels(3).Text = CStr(Datos.RecordCount) & " registros"
    MdiExport.ExportarDatos Datos, "INFORME DE CLEVER SOFTPRO", Ruta
    Datos.Close
    Exit Sub
ErrExportar:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    On Error GoTo 0
    Err.Raise Numero, , Descripcion
End Sub

Private Sub CmdgenExcel_Click()
    On Error GoTo ErrInt
    ExportarInforme
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub Command1_Click()
    Dim Datos As ADODB.Recordset
    On Error GoTo ErrInt
    Set Datos = GenerarInforme()
    If Datos.RecordCount = 0 Then Err.Raise 5, , "El informe no contiene registros."
    CUAD.CancelError = True: CUAD.FileName = "": CUAD.DialogTitle = "Guardar Informe"
    CUAD.Filter = "Informe de texto (*.txt)|*.txt": CUAD.DefaultExt = "txt"
    CUAD.Flags = cdlOFNOverwritePrompt Or cdlOFNPathMustExist
    CUAD.ShowSave
    GuardarInformeTexto Datos, CUAD.FileName, True
    Datos.Close
    Exit Sub
ErrInt:
    Dim Numero As Long, Descripcion As String
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    On Error GoTo 0
    If Numero <> cdlCancel Then MsgBox Descripcion, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdImprimir_Click()
    Dim Datos As ADODB.Recordset
    On Error GoTo ErrInt
    Set Datos = GenerarInforme()
    If Datos.RecordCount = 0 Then Err.Raise 5, , "El informe no contiene registros."
    If MsgBox("Se imprimiran " & CStr(Datos.RecordCount) & " registros. Desea continuar?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") = vbYes Then ImprimirInforme Datos, "INFORME DE CLEVER SOFTPRO"
    Datos.Close
    Exit Sub
ErrInt:
    Dim Descripcion As String
    Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    On Error GoTo 0
    MsgBox Descripcion, vbExclamation, "Clever SoftPro"
End Sub
Private Sub ActualizarAcciones()
    Dim Nodo As MSComctlLib.Node, Cantidad As Long
    For Each Nodo In ViewMenus.Nodes
        If Nodo.Checked Then Cantidad = Cantidad + 1
    Next
    CmdgenExcel.Enabled = (Cantidad > 0): Command1.Enabled = (Cantidad > 0): CmdImprimir.Enabled = (Cantidad > 0)
    If CmbOrigen.ListIndex >= 0 Then StatusBar1.Panels(1).Text = CmbOrigen.Text Else StatusBar1.Panels(1).Text = "Seleccione el origen del informe"
    StatusBar1.Panels(2).Text = CStr(Cantidad) & " columnas | Filtros Y/O de izquierda a derecha"
    StatusBar1.Panels(3).Text = "Consulta pendiente"
End Sub
Private Sub TxtFil_Change(Index As Integer)
    If Not Preparando Then ActualizarAcciones
End Sub
Private Sub OptY_Click(Index As Integer)
    If Not Preparando Then ActualizarAcciones
End Sub
Private Sub OptO_Click(Index As Integer)
    If Not Preparando Then ActualizarAcciones
End Sub
Private Sub Form_Activate()
    VerificarAccesoFormulario Me, 9
End Sub
