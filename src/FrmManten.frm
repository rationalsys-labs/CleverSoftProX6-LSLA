VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FrmManten 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Clientes"
   ClientHeight    =   7380
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   15045
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   6750
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManten.frx":0
            Key             =   "Crear"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManten.frx":39A
            Key             =   "Modificar"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManten.frx":934
            Key             =   "Eliminar"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManten.frx":CCE
            Key             =   "Imprimir"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar BarHer 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   635
      ButtonWidth     =   2858
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Crear Cliente"
            Key             =   "Crear"
            ImageKey        =   "Crear"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Modificar Cliente"
            Key             =   "Modificar"
            ImageKey        =   "Modificar"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Eliminar Cliente"
            Key             =   "Eliminar"
            ImageKey        =   "Eliminar"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir Registro"
            Key             =   "Imprimir"
            ImageKey        =   "Imprimir"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame FraGen 
      Height          =   6630
      Left            =   90
      TabIndex        =   33
      Top             =   405
      Width           =   4155
      Begin VB.TextBox TxtBusCli 
         Height          =   285
         Left            =   90
         TabIndex        =   37
         Top             =   180
         Width           =   3930
      End
      Begin MSFlexGridLib.MSFlexGrid GridClientes 
         Height          =   5955
         Left            =   90
         TabIndex        =   34
         Top             =   540
         Width           =   3930
         _ExtentX        =   6932
         _ExtentY        =   10504
         _Version        =   393216
         FixedCols       =   0
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   35
      Top             =   7140
      Width           =   15045
      _ExtentX        =   26538
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3704
            MinWidth        =   3704
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
   Begin VB.Frame FraInf 
      Height          =   6630
      Left            =   4275
      TabIndex        =   0
      Top             =   405
      Width           =   10680
      Begin VB.TextBox TxtCli 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   2160
         TabIndex        =   17
         Top             =   270
         Width           =   4965
      End
      Begin VB.TextBox TxtDir 
         BackColor       =   &H00EAF5FF&
         Height          =   315
         Left            =   2160
         TabIndex        =   16
         Top             =   675
         Width           =   4965
      End
      Begin VB.TextBox TxtDoc 
         BackColor       =   &H00EAF5FF&
         Height          =   315
         Left            =   4410
         TabIndex        =   15
         Top             =   1080
         Width           =   2715
      End
      Begin VB.ComboBox CmbDoc 
         BackColor       =   &H00EAF5FF&
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1080
         Width           =   2220
      End
      Begin VB.TextBox TxtCod 
         BackColor       =   &H00EAF5FF&
         Height          =   315
         Left            =   8460
         TabIndex        =   13
         Top             =   270
         Width           =   1995
      End
      Begin VB.ComboBox CmbDis 
         BackColor       =   &H00EAF5FF&
         Height          =   315
         Left            =   8460
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   675
         Width           =   1995
      End
      Begin VB.TextBox TxtRep 
         BackColor       =   &H00EAF5FF&
         Height          =   315
         Left            =   2160
         TabIndex        =   10
         Top             =   2070
         Width           =   4965
      End
      Begin VB.TextBox TxtMailRep 
         BackColor       =   &H00EAF5FF&
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   2475
         Width           =   4965
      End
      Begin VB.TextBox TxtCont 
         BackColor       =   &H00EAF5FF&
         Height          =   315
         Left            =   2160
         TabIndex        =   8
         Top             =   2880
         Width           =   4965
      End
      Begin VB.TextBox TxtMailCont 
         BackColor       =   &H00EAF5FF&
         Height          =   315
         Left            =   2160
         TabIndex        =   7
         Top             =   3285
         Width           =   4965
      End
      Begin VB.TextBox TxtTelGen 
         BackColor       =   &H00EAF5FF&
         Height          =   315
         Left            =   2160
         TabIndex        =   6
         Top             =   3690
         Width           =   2220
      End
      Begin VB.TextBox TxtTeltra 
         BackColor       =   &H00EAF5FF&
         Height          =   315
         Left            =   2160
         TabIndex        =   5
         Top             =   4095
         Width           =   2220
      End
      Begin VB.TextBox TxtTelMov 
         BackColor       =   &H00EAF5FF&
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   4500
         Width           =   2220
      End
      Begin VB.ComboBox CmbEst 
         BackColor       =   &H00EAF5FF&
         Height          =   315
         Left            =   2160
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   5310
         Width           =   1500
      End
      Begin VB.TextBox TxtTelFax 
         BackColor       =   &H00EAF5FF&
         Height          =   315
         Left            =   2160
         TabIndex        =   2
         Top             =   4905
         Width           =   2220
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8550
         TabIndex        =   38
         Top             =   5670
         Width           =   1995
      End
      Begin VB.CommandButton CmdGuardar 
         Caption         =   "Guardar Cambios"
         Enabled         =   0   'False
         Height          =   375
         Left            =   8550
         TabIndex        =   1
         Top             =   6120
         Width           =   1995
      End
      Begin MSComCtl2.DTPicker DtpFec 
         Height          =   315
         Left            =   2160
         TabIndex        =   11
         Top             =   1485
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   17956865
         CurrentDate     =   40461
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   225
         TabIndex        =   32
         Top             =   270
         Width           =   525
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
         Height          =   195
         Left            =   225
         TabIndex        =   31
         Top             =   675
         Width           =   720
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Documento de Identidad:"
         Height          =   195
         Left            =   225
         TabIndex        =   30
         Top             =   1080
         Width           =   1800
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   7425
         TabIndex        =   29
         Top             =   315
         Width           =   540
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Distrito:"
         Height          =   195
         Left            =   7425
         TabIndex        =   28
         Top             =   720
         Width           =   525
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Apertura:"
         Height          =   195
         Left            =   225
         TabIndex        =   27
         Top             =   1485
         Width           =   1365
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Representante Legal:"
         Height          =   195
         Left            =   225
         TabIndex        =   26
         Top             =   2115
         Width           =   1530
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Email:"
         Height          =   195
         Left            =   225
         TabIndex        =   25
         Top             =   2475
         Width           =   420
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Contacto Empresa:"
         Height          =   195
         Left            =   225
         TabIndex        =   24
         Top             =   2925
         Width           =   1350
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Email:"
         Height          =   195
         Left            =   225
         TabIndex        =   23
         Top             =   3285
         Width           =   420
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Telefono General:"
         Height          =   195
         Left            =   225
         TabIndex        =   22
         Top             =   3690
         Width           =   1275
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Telefono de Trabajo:"
         Height          =   195
         Left            =   225
         TabIndex        =   21
         Top             =   4095
         Width           =   1485
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Telefono Movil:"
         Height          =   195
         Left            =   225
         TabIndex        =   20
         Top             =   4500
         Width           =   1095
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Estado de Cliente:"
         Height          =   195
         Left            =   225
         TabIndex        =   19
         Top             =   5310
         Width           =   1290
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Telefono de Fax:"
         Height          =   195
         Left            =   225
         TabIndex        =   18
         Top             =   4905
         Width           =   1200
      End
   End
End
Attribute VB_Name = "FrmManten"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private EnEdicion As Boolean
Private TipoGuardado As Integer
Private DatosClientes As ADODB.Recordset
Private Cargando As Boolean
Private CodigoAnteriorEdicion As String

Private Sub Form_Load()
    On Error GoTo ErrInt
    EnEdicion = False: Cargando = False
    StatusBar1.Panels(1).Width = 2400
    StatusBar1.Panels(2).AutoSize = sbrSpring
    StatusBar1.Panels(3).Width = 2500
    Me.Icon = MDIPri.ICO.ListImages(25).Picture
    GridClientes.Rows = 1: GridClientes.Cols = 16
    Dim Columna As Integer
    For Columna = 0 To 15: GridClientes.ColWidth(Columna) = 0: Next
    GridClientes.ColWidth(0) = 3600
    GridClientes.TextMatrix(0, 0) = "Nombre del Cliente"
    CmbEst.AddItem "Activo": CmbEst.ItemData(0) = 1
    CmbEst.AddItem "Inactivo": CmbEst.ItemData(1) = 0
    CargarCatalogo CmbDis, "SELECT CoDist,NomDis FROM dbo.MDBDistritos ORDER BY NomDis"
    CargarCatalogo CmbDoc, "SELECT CodTip,DesTip FROM dbo.IntTipDocument ORDER BY CodTip"
    TxtCli.MaxLength = 300: TxtDir.MaxLength = 300: TxtDoc.MaxLength = 20
    TxtRep.MaxLength = 150: TxtMailRep.MaxLength = 150: TxtCont.MaxLength = 150: TxtMailCont.MaxLength = 150
    TxtTelGen.MaxLength = 50: TxtTeltra.MaxLength = 50: TxtTelMov.MaxLength = 50: TxtTelFax.MaxLength = 50
    RecargarClientes
    EstadoEdicion False
    Exit Sub
ErrInt:
    InformarError
End Sub

Private Sub CargarCatalogo(ByVal Lista As ComboBox, ByVal Consulta As String)
    Dim Datos As ADODB.Recordset
    Set Datos = ConsultarDatos(Consulta)
    Lista.Clear
    Do Until Datos.EOF
        Lista.AddItem CStr(Datos.Fields(1).Value)
        Lista.ItemData(Lista.NewIndex) = CLng(Datos.Fields(0).Value)
        Datos.MoveNext
    Loop
    Datos.Close
End Sub

Public Sub RecargarClientes()
    Dim NuevosDatos As ADODB.Recordset
    Set NuevosDatos = ConsultarDatos("SELECT v.Cliente,v.Codigo,v.Direccion,v.Distrito,v.Tipo,v.Documento,v.Apertura,v.Representante,v.Mail,v.Contacto,v.MailContacto,v.Telefono,v.TelTrabajo,v.TelMovil,v.TelFax,v.Estado,c.TipDoc,c.DisCli FROM dbo.ViewCliente v JOIN dbo.MDBCostumer c ON c.CodCli=v.Codigo ORDER BY v.Cliente,v.Codigo")
    If Not DatosClientes Is Nothing Then
        If DatosClientes.State <> adStateClosed Then DatosClientes.Close
    End If
    Set DatosClientes = NuevosDatos
    FiltrarClientes
End Sub

Private Sub FiltrarClientes()
    Dim Fila As Long
    Dim Columna As Integer
    Cargando = True
    GridClientes.Redraw = False
    GridClientes.Rows = 1
    If DatosClientes.RecordCount > 0 Then DatosClientes.MoveFirst
    Do Until DatosClientes.EOF
        If InStr(1, TextoDato(DatosClientes.Fields(0).Value), TxtBusCli.Text, vbTextCompare) > 0 Or Len(TxtBusCli.Text) = 0 Then
            GridClientes.Rows = GridClientes.Rows + 1
            Fila = GridClientes.Rows - 1
            For Columna = 0 To 15
                GridClientes.TextMatrix(Fila, Columna) = TextoDato(DatosClientes.Fields(Columna).Value)
            Next
        End If
        DatosClientes.MoveNext
    Loop
    GridClientes.Redraw = True
    Cargando = False
    If Not EnEdicion Then LimpiarInf
    FraGen.Visible = True: FraInf.Visible = True
    ActualizarEstado
End Sub

Private Sub TxtBusCli_Change()
    On Error GoTo ErrInt
    If Cargando Or DatosClientes Is Nothing Then Exit Sub
    FiltrarClientes
    Exit Sub
ErrInt:
    GridClientes.Redraw = True: Cargando = False
    InformarError
End Sub

Private Function PuedeAbandonar() As Boolean
    PuedeAbandonar = True
    If EnEdicion Then PuedeAbandonar = (MsgBox("Hay cambios sin guardar. Desea descartarlos?", vbQuestion Or vbYesNo, "Clever SoftPro") = vbYes)
End Function

Private Sub EstadoEdicion(ByVal Editando As Boolean)
    Dim Control As Control
    Dim Boton As Integer
    EnEdicion = Editando
    If Not Editando Then TipoGuardado = 0
    For Each Control In Me.Controls
        If TypeOf Control Is TextBox Then
            If Control.Name <> "TxtBusCli" Then
                Control.Locked = Not Editando
                If Editando Then Control.BackColor = &HFFFFFF Else Control.BackColor = &HEAF5FF
            End If
        End If
    Next
    TxtCod.Locked = True
    CmbDis.Enabled = Editando: CmbDoc.Enabled = Editando: CmbEst.Enabled = Editando
    DtpFec.Enabled = False: CmdGuardar.Enabled = Editando: CmdCancelar.Enabled = Editando
    BarHer.Enabled = Not Editando
    For Boton = 1 To BarHer.Buttons.Count
        BarHer.Buttons(Boton).Enabled = Not Editando
    Next Boton
    BarHer.Refresh
    ActualizarEstado
End Sub

Private Sub ActualizarEstado(Optional ByVal Mensaje As String = "")
    If Len(Mensaje) = 0 Then
        If EnEdicion Then
            If TipoGuardado = 1 Then Mensaje = "Creando cliente" Else Mensaje = "Modificando cliente"
        Else
            Mensaje = "Consulta"
        End If
    End If
    StatusBar1.Panels(1).Text = Mensaje
    If EnEdicion And TipoGuardado = 1 Then
        StatusBar1.Panels(2).Text = "Nuevo cliente"
    ElseIf Len(TxtCod.Text) > 0 Then
        StatusBar1.Panels(2).Text = TxtCod.Text & " - " & TxtCli.Text
    Else
        StatusBar1.Panels(2).Text = "Seleccione un cliente de la lista"
    End If
    If Not DatosClientes Is Nothing Then
        If DatosClientes.State <> adStateClosed Then StatusBar1.Panels(3).Text = "Mostrados: " & CStr(GridClientes.Rows - 1) & " / " & CStr(DatosClientes.RecordCount)
    End If
End Sub

Private Sub LimpiarInf()
    Dim Control As Control
    For Each Control In Me.Controls
        If TypeOf Control Is TextBox Then
            If Control.Name <> "TxtBusCli" Then Control.Text = vbNullString
        End If
    Next
    CmbDis.ListIndex = -1: CmbDoc.ListIndex = -1: CmbEst.ListIndex = -1
    DtpFec.Value = Date
    ActualizarEstado
End Sub

Private Sub SeleccionarCodigo(ByVal Lista As ComboBox, ByVal Codigo As Long)
    Dim Posicion As Long
    Lista.ListIndex = -1
    For Posicion = 0 To Lista.ListCount - 1
        If Lista.ItemData(Posicion) = Codigo Then Lista.ListIndex = Posicion: Exit For
    Next
End Sub

Public Sub MostrarCliente(ByVal Codigo As String)
    DatosClientes.MoveFirst
    DatosClientes.Find "Codigo = " & CStr(CDec(Codigo))
    If DatosClientes.EOF Then Err.Raise vbObjectError + 203, , "El cliente ya no esta disponible."
    LimpiarInf
    TxtCli.Text = TextoDato(DatosClientes.Fields(0).Value): TxtCod.Text = CStr(DatosClientes.Fields(1).Value)
    TxtDir.Text = TextoDato(DatosClientes.Fields(2).Value): TxtDoc.Text = TextoDato(DatosClientes.Fields(5).Value)
    DtpFec.Value = DatosClientes.Fields(6).Value
    TxtRep.Text = TextoDato(DatosClientes.Fields(7).Value): TxtMailRep.Text = TextoDato(DatosClientes.Fields(8).Value)
    TxtCont.Text = TextoDato(DatosClientes.Fields(9).Value): TxtMailCont.Text = TextoDato(DatosClientes.Fields(10).Value)
    TxtTelGen.Text = TextoDato(DatosClientes.Fields(11).Value): TxtTeltra.Text = TextoDato(DatosClientes.Fields(12).Value)
    TxtTelMov.Text = TextoDato(DatosClientes.Fields(13).Value): TxtTelFax.Text = TextoDato(DatosClientes.Fields(14).Value)
    SeleccionarCodigo CmbDoc, CLng(DatosClientes.Fields(16).Value)
    SeleccionarCodigo CmbDis, CLng(DatosClientes.Fields(17).Value)
    If CStr(DatosClientes.Fields(15).Value) = "Activo" Then CmbEst.ListIndex = 0 Else CmbEst.ListIndex = 1
    FraGen.Visible = True: FraInf.Visible = True
    ActualizarEstado
End Sub

Private Sub RestaurarSeleccion()
    Dim Fila As Long, Codigo As String
    Codigo = TxtCod.Text
    If EnEdicion And TipoGuardado = 1 Then Codigo = CodigoAnteriorEdicion
    Cargando = True
    GridClientes.Row = 0: GridClientes.RowSel = 0
    For Fila = 1 To GridClientes.Rows - 1
        If GridClientes.TextMatrix(Fila, 1) = Codigo Then
            GridClientes.Row = Fila: GridClientes.RowSel = Fila
            Exit For
        End If
    Next
    Cargando = False
End Sub

Private Sub GridClientes_Click()
    If Cargando Or GridClientes.Row < 1 Then Exit Sub
    On Error GoTo ErrInt
    If Not PuedeAbandonar() Then
        RestaurarSeleccion
        Exit Sub
    End If
    EstadoEdicion False
    MostrarCliente GridClientes.TextMatrix(GridClientes.Row, 1)
    Exit Sub
ErrInt:
    InformarError
End Sub

Private Sub CmdNewCli_Click()
    Dim Datos As ADODB.Recordset
    On Error GoTo ErrInt
    If Not PuedeAbandonar() Then Exit Sub
    CodigoAnteriorEdicion = TxtCod.Text
    LimpiarInf
    Set Datos = ConsultarDatos("EXEC dbo.Correlativo 'CLI'")
    TxtCod.Text = CStr(Datos.Fields(0).Value): Datos.Close
    TipoGuardado = 1: EstadoEdicion True: CmbEst.ListIndex = 0: CmbEst.Enabled = False
    FraGen.Visible = True: FraInf.Visible = True
    TxtCli.SetFocus
    Exit Sub
ErrInt:
    InformarError
End Sub

Private Sub CmdModCli_Click()
    If EnEdicion Then Exit Sub
    If Len(TxtCod.Text) = 0 Then
        MsgBox "Seleccione un cliente de la lista antes de modificarlo.", vbInformation, "Modificar Cliente"
        Exit Sub
    End If
    CodigoAnteriorEdicion = TxtCod.Text
    TipoGuardado = 2: EstadoEdicion True
    TxtCli.SetFocus
End Sub

Private Sub CmdCancelar_Click()
    On Error GoTo ErrInt
    If Not EnEdicion Then Exit Sub
    If Not PuedeAbandonar() Then Exit Sub
    EstadoEdicion False
    TipoGuardado = 0
    If Len(CodigoAnteriorEdicion) > 0 Then
        MostrarCliente CodigoAnteriorEdicion
    Else
        LimpiarInf
    End If
    CodigoAnteriorEdicion = vbNullString
    ActualizarEstado "Cambios cancelados"
    Exit Sub
ErrInt:
    InformarError
End Sub

Private Function DatosCompletos() As Boolean
    Dim Faltantes As String
    If Len(Trim$(TxtCli.Text)) = 0 Then Faltantes = Faltantes & vbCrLf & "- Nombre del cliente"
    If Len(Trim$(TxtDir.Text)) = 0 Then Faltantes = Faltantes & vbCrLf & "- Direccion"
    If CmbDoc.ListIndex < 0 Then Faltantes = Faltantes & vbCrLf & "- Tipo de documento"
    If Len(Trim$(TxtDoc.Text)) = 0 Then Faltantes = Faltantes & vbCrLf & "- Numero de documento"
    If CmbDis.ListIndex < 0 Then Faltantes = Faltantes & vbCrLf & "- Distrito"
    If Len(Trim$(TxtRep.Text)) = 0 Then Faltantes = Faltantes & vbCrLf & "- Representante legal"
    If Len(Trim$(TxtCont.Text)) = 0 Then Faltantes = Faltantes & vbCrLf & "- Contacto"
    If Len(Trim$(TxtTelGen.Text)) = 0 Then Faltantes = Faltantes & vbCrLf & "- Telefono general"
    If CmbEst.ListIndex < 0 Then Faltantes = Faltantes & vbCrLf & "- Estado"
    DatosCompletos = (Len(Faltantes) = 0)
    If Not DatosCompletos Then MsgBox "Complete los siguientes campos obligatorios:" & vbCrLf & Faltantes, vbInformation, "Datos del Cliente"
End Function

Private Sub CmdGuardar_Click()
    Dim Cliente As DatosCliente
    Dim Codigo As String
    Dim Guardado As Boolean, Descripcion As String
    If Not EnEdicion Then Exit Sub
    If Not DatosCompletos() Then Exit Sub
    On Error GoTo ErrInt
    CmdGuardar.Enabled = False
    Cliente.Codigo = TxtCod.Text: Cliente.Nombre = Trim$(TxtCli.Text)
    Cliente.Documento = Trim$(TxtDoc.Text): Cliente.TipoDocumento = CmbDoc.ItemData(CmbDoc.ListIndex)
    Cliente.Direccion = Trim$(TxtDir.Text): Cliente.Distrito = CmbDis.ItemData(CmbDis.ListIndex)
    Cliente.Representante = TxtRep.Text: Cliente.CorreoRepresentante = TxtMailRep.Text
    Cliente.Contacto = TxtCont.Text: Cliente.CorreoContacto = TxtMailCont.Text
    Cliente.Telefono = TxtTelGen.Text: Cliente.TelefonoTrabajo = TxtTeltra.Text
    Cliente.TelefonoMovil = TxtTelMov.Text: Cliente.TelefonoFax = TxtTelFax.Text
    Cliente.Activo = (CmbEst.ItemData(CmbEst.ListIndex) = 1)
    Codigo = GuardarCliente(Cliente, TipoGuardado)
    Guardado = True: TxtCod.Text = Codigo: CodigoAnteriorEdicion = vbNullString
    EstadoEdicion False
    RecargarClientes
    MostrarCliente Codigo
    RestaurarSeleccion
    ActualizarEstado "Cliente guardado"
    Exit Sub
ErrInt:
    Descripcion = Err.Description
    If Guardado Then
        EstadoEdicion False
        TxtCod.Text = Codigo
        ActualizarEstado "Cliente guardado; pendiente recargar"
        MsgBox "El cliente con codigo " & Codigo & " fue guardado correctamente, pero no se pudo actualizar la consulta: " & Descripcion & vbCrLf & "Vuelva a abrir la consulta para recargar los datos. No repita el guardado.", vbExclamation, "Clever SoftPro"
    Else
        CmdGuardar.Enabled = EnEdicion
        InformarError
    End If
End Sub

Private Sub CmdEliCli_Click()
    On Error GoTo ErrInt
    If EnEdicion Then Exit Sub
    If Len(TxtCod.Text) = 0 Then
        MsgBox "Seleccione un cliente de la lista antes de eliminarlo.", vbInformation, "Eliminar Cliente"
        Exit Sub
    End If
    If MsgBox("Se dara de baja al siguiente cliente:" & vbCrLf & vbCrLf & _
        "Codigo: " & TxtCod.Text & vbCrLf & _
        "Cliente: " & TxtCli.Text & vbCrLf & _
        "Documento: " & CmbDoc.Text & " - " & TxtDoc.Text & vbCrLf & vbCrLf & _
        "Se conservaran sus registros e historial." & vbCrLf & _
        "Desea continuar?", vbQuestion Or vbYesNo Or vbDefaultButton2, "Eliminar Cliente") <> vbYes Then Exit Sub
    EliminarCliente TxtCod.Text
    LimpiarInf
    RecargarClientes
    ActualizarEstado "Cliente dado de baja"
    Exit Sub
ErrInt:
    InformarError
End Sub

Private Sub CmdGen_Click()
    If Not PuedeAbandonar() Then Exit Sub
    EstadoEdicion False
    FraGen.Visible = True: FraInf.Visible = True
    ActualizarEstado
End Sub

Private Sub BarHer_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Crear": CmdNewCli_Click
        Case "Modificar": CmdModCli_Click
        Case "Eliminar": CmdEliCli_Click
        Case Else
            If Button.Index = 4 Then ImprimirCliente
    End Select
End Sub

Private Sub ImprimirCliente()
    On Error GoTo ErrInt
    If Len(TxtCod.Text) = 0 Then Exit Sub
    If MsgBox("Desea imprimir el registro del cliente?", vbQuestion Or vbYesNo, "Imprimir Registro") <> vbYes Then Exit Sub
    Printer.FontName = "Arial": Printer.FontSize = 10
    Printer.Print "Clever SoftPro - Registro de cliente"
    Printer.Print "Codigo: " & TxtCod.Text
    Printer.Print "Cliente: " & TxtCli.Text
    Printer.Print "Documento: " & TxtDoc.Text
    Printer.Print "Direccion: " & TxtDir.Text
    Printer.Print "Distrito: " & CmbDis.Text
    Printer.Print "Contacto: " & TxtCont.Text
    Printer.Print "Telefono: " & TxtTelGen.Text
    Printer.EndDoc
    Exit Sub
ErrInt:
    InformarError
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not CerrandoAplicacion Then Cancel = Not PuedeAbandonar()
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not DatosClientes Is Nothing Then
        If DatosClientes.State <> adStateClosed Then DatosClientes.Close
    End If
End Sub

Private Sub InformarError()
    MsgBox "No se pudo completar la operacion: " & Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Public Function PuedeCerrar() As Boolean
    PuedeCerrar = PuedeAbandonar()
End Function

Private Sub Form_Activate()
    VerificarAccesoFormulario Me, 1
End Sub
