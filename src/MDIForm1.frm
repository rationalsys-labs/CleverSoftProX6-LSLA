VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F5E116E1-0563-11D8-AA80-000B6A0D10CB}#1.0#0"; "HookMenu.ocx"
Begin VB.MDIForm MDIPri 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Clever SoftPro"
   ClientHeight    =   9600
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   14130
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer TmrCerrar 
      Enabled = 0
      Interval = 1
      Left = 120
      Top = 120
   End
   Begin MSComctlLib.Toolbar Barher 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   14130
      _ExtentX        =   24924
      _ExtentY        =   635
      ButtonWidth     =   1905
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      TextAlignment   =   1
      ImageList       =   "ICO"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   5
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Registros"
            Key             =   "Registros"
            ImageKey        =   "a26"
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Clientes  "
            Key             =   "Clientes"
            ImageKey        =   "A30"
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Etiquetas"
            Key             =   "Etiquetas"
            ImageKey        =   "A29"
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cargos   "
            Key             =   "Cargos"
            ImageKey        =   "A27"
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Status   "
            Key             =   "Status"
            ImageKey        =   "A28"
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin HookMenu.XpMenu XpMenu1 
      Left            =   1890
      Top             =   4770
      _ExtentX        =   900
      _ExtentY        =   900
      BmpCount        =   8
      CheckBorderColor=   7021576
      SelMenuBorder   =   7021576
      SelMenuBackColor=   14073525
      SelMenuForeColor=   16646297
      SelCheckBackColor=   13740436
      MenuBorderColor =   6956042
      SeparatorColor  =   -2147483632
      MenuBackColor   =   14609903
      MenuForeColor   =   0
      CheckBackColor  =   15326939
      CheckForeColor  =   10027263
      DisabledMenuBorderColor=   -2147483632
      DisabledMenuBackColor=   15660791
      DisabledMenuForeColor=   -2147483631
      MenuBarBackColor=   14215660
      MenuPopupBackColor=   16777215
      ShortCutNormalColor=   0
      ShortCutSelectColor=   16646297
      ArrowNormalColor=   10027263
      ArrowSelectColor=   12484864
      ShadowColor     =   0
      Bmp:1           =   "MDIForm1.frx":038A
      Mask:1          =   14876671
      Key:1           =   "#OpeVis"
      Bmp:2           =   "MDIForm1.frx":08DC
      Mask:2          =   15138815
      Key:2           =   "#Mantenimiento"
      Bmp:3           =   "MDIForm1.frx":0E2E
      Key:3           =   "#Registros"
      Bmp:4           =   "MDIForm1.frx":1256
      Key:4           =   "#Reparto"
      Bmp:5           =   "MDIForm1.frx":167E
      Key:5           =   "#RegStatus"
      Bmp:6           =   "MDIForm1.frx":1AA6
      Key:6           =   "#OpeGrup"
      Bmp:7           =   "MDIForm1.frx":1ECE
      Key:7           =   "#Consol"
      Bmp:8           =   "MDIForm1.frx":22F6
      Key:8           =   "#Cntdeli"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      Align           =   4  'Align Right
      Height          =   9000
      Left            =   12225
      ScaleHeight     =   8940
      ScaleWidth      =   1845
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1905
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   420
         Left            =   540
         TabIndex        =   6
         Top             =   3555
         Visible         =   0   'False
         Width           =   1230
      End
      Begin VB.ListBox LstConf 
         Height          =   2400
         Left            =   135
         TabIndex        =   2
         Top             =   225
         Width           =   1500
      End
      Begin VB.PictureBox PicFondo 
         Height          =   1770
         Left            =   180
         Picture         =   "MDIForm1.frx":271E
         ScaleHeight     =   1710
         ScaleWidth      =   1470
         TabIndex        =   1
         Top             =   5130
         Width           =   1530
      End
      Begin MSComctlLib.ProgressBar PROGRESS 
         Height          =   255
         Left            =   180
         TabIndex        =   4
         Top             =   4770
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin MSComctlLib.ImageList ICO 
      Left            =   765
      Top             =   4815
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   31
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1C249
            Key             =   "Cli"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1C7E3
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1D1F5
            Key             =   "a3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1DACF
            Key             =   "a4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1DD64
            Key             =   "a5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1E2FE
            Key             =   "a6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1E898
            Key             =   "a7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1EE32
            Key             =   "a8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F3CC
            Key             =   "a9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1F966
            Key             =   "a10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1FF00
            Key             =   "a11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2049A
            Key             =   "a12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":20A34
            Key             =   "a13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":20FCE
            Key             =   "a14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":21568
            Key             =   "a15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":21B02
            Key             =   "a16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2209C
            Key             =   "a17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":22636
            Key             =   "a18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":22BD0
            Key             =   "a19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2316A
            Key             =   "a20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":23704
            Key             =   "a21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":23C9E
            Key             =   "a22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":24238
            Key             =   "a23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":247D2
            Key             =   "a24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":251E4
            Key             =   "a25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2577E
            Key             =   "a26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":25B8A
            Key             =   "A27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":26124
            Key             =   "A28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":266BE
            Key             =   "A29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":26C58
            Key             =   "A30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":27080
            Key             =   "a31"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   9360
      Width           =   14130
      _ExtentX        =   24924
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
   Begin VB.Menu Archivo 
      Caption         =   "Archivo"
      Begin VB.Menu Salir 
         Caption         =   "Salir"
      End
   End
   Begin VB.Menu Operaciones 
      Caption         =   "Operaciones"
      Begin VB.Menu OpeVis 
         Caption         =   "Visualizar Informacion de Registros"
      End
      Begin VB.Menu Mantenimiento 
         Caption         =   "Mantenimiento de Clientes"
      End
      Begin VB.Menu Registros 
         Caption         =   "Crear Registros de Reparto"
      End
      Begin VB.Menu OpeConRep 
         Caption         =   "Control de Datos de Reparto"
         Begin VB.Menu Reparto 
            Caption         =   "Contro de Reparto y Ranches"
         End
         Begin VB.Menu Consol 
            Caption         =   "Generar Consolidado de Reparto"
         End
         Begin VB.Menu RegStatus 
            Caption         =   "Registrar Status"
         End
         Begin VB.Menu Cntdeli 
            Caption         =   "Control de Delivery"
         End
      End
      Begin VB.Menu OpeGrup 
         Caption         =   "Gestionar Grupos de Servicios"
      End
   End
   Begin VB.Menu Forinf 
      Caption         =   "Formularios e Informes"
      Begin VB.Menu ForStatus 
         Caption         =   "Informe de Status"
      End
   End
   Begin VB.Menu Configuracion 
      Caption         =   "Configuracion"
      Begin VB.Menu ConfVar 
         Caption         =   "Configurar Variables de Sistema"
      End
      Begin VB.Menu confActu 
         Caption         =   "Actualizaciones"
      End
   End
   Begin VB.Menu Usuarios
      Caption = "Usuarios"
      Begin VB.Menu AdministrarUsuarios
         Caption = "Administracion de Usuarios"
      End
   End
   Begin VB.Menu Ayuda 
      Caption         =   "Ayuda"
      Begin VB.Menu ayuCon 
         Caption         =   "Contenido"
      End
      Begin VB.Menu AyuAce 
         Caption         =   "Acerca de Clever SoftPro X6"
      End
   End
   Begin VB.Menu MenuBus 
      Caption         =   "Menubusqueda"
      Visible         =   0   'False
      Begin VB.Menu BusReg 
         Caption         =   "Buscar Registros"
      End
      Begin VB.Menu busEli 
         Caption         =   "Eliminar filtro"
      End
   End
End
Attribute VB_Name = "MDIPri"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Const TITULO As String = "Clever SoftPro"

Private Sub MDIForm_Load()
    On Error GoTo ErrInt
    Operaciones.Enabled = False
    Forinf.Enabled = False
    Configuracion.Enabled = False
    Barher.Enabled = False
    SetParent PROGRESS.hWnd, StatusBar1.hWnd
    AjustarProgreso
    Me.Picture = PicFondo.Picture
    ActualizarSesion
    Exit Sub
ErrInt:
    Err.Raise Err.Number, "MDIPri.MDIForm_Load", Err.Description
End Sub

Private Sub AjustarProgreso()
    If Me.WindowState = vbMinimized Then Exit Sub
    PROGRESS.Move StatusBar1.Panels(2).Left, 35, StatusBar1.Panels(2).Width, StatusBar1.Height - 40
End Sub

Public Sub CentrarLogin()
    Dim Formulario As Object
    If CerrandoAplicacion Or EjecuteDatrav Then Exit Sub
    'No acceder a la instancia predeterminada aqui: podria cargar el login de nuevo.
    For Each Formulario In Forms
        If TypeOf Formulario Is FrmLogin Then
            If Me.ScaleWidth > Formulario.Width Then Formulario.Left = (Me.ScaleWidth - Formulario.Width) / 2
            If Me.ScaleHeight > Formulario.Height Then Formulario.Top = (Me.ScaleHeight - Formulario.Height) / 2
            Exit For
        End If
    Next Formulario
End Sub

Public Sub ActualizarSesion()
    Me.Caption = TITULO
    If EjecuteDatrav Then
        Me.Caption = TITULO & " " & Format$(Date, "dd/mm/yyyy") & " - Conectado bajo Usuario " & UCase$(VarUsuario)
        StatusBar1.Panels(1).Text = VarUsuario
        StatusBar1.Panels(3).Text = "Sesion iniciada"
    Else
        StatusBar1.Panels(1).Text = vbNullString
        StatusBar1.Panels(3).Text = "Pendiente de autenticacion"
    End If
    Forinf.Enabled = EjecuteDatrav: Configuracion.Enabled = EjecuteDatrav
    Operaciones.Enabled = EjecuteDatrav
    OpeVis.Enabled = EjecuteDatrav: Registros.Enabled = EjecuteDatrav: OpeConRep.Enabled = EjecuteDatrav: OpeGrup.Enabled = EjecuteDatrav
    Reparto.Enabled = EjecuteDatrav: Consol.Enabled = EjecuteDatrav
    RegStatus.Enabled = EjecuteDatrav: Cntdeli.Enabled = EjecuteDatrav
    Mantenimiento.Enabled = EjecuteDatrav
    Barher.Enabled = EjecuteDatrav
    Dim Boton As Integer
    For Boton = 1 To Barher.Buttons.Count
        Barher.Buttons(Boton).Enabled = (EjecuteDatrav And (Barher.Buttons(Boton).Key = "Clientes" Or Barher.Buttons(Boton).Key = "Registros" Or Barher.Buttons(Boton).Key = "Etiquetas" Or Barher.Buttons(Boton).Key = "Cargos" Or Barher.Buttons(Boton).Key = "Status"))
    Next

    Usuarios.Enabled = TieneAcceso(0)
    Mantenimiento.Enabled = TieneAcceso(1)
    OpeVis.Enabled = TieneAcceso(2)
    Registros.Enabled = TieneAcceso(3)
    OpeGrup.Enabled = TieneAcceso(4)
    Reparto.Enabled = TieneAcceso(5)
    Consol.Enabled = TieneAcceso(6)
    RegStatus.Enabled = TieneAcceso(7)
    Cntdeli.Enabled = TieneAcceso(8)
    ForStatus.Enabled = TieneAcceso(9)
    ConfVar.Enabled = TieneAcceso(10)
    confActu.Enabled = TieneAcceso(11)
    Barher.Buttons("Clientes").Enabled = TieneAcceso(1)
    Barher.Buttons("Registros").Enabled = TieneAcceso(2)
    Barher.Buttons("Etiquetas").Enabled = TieneAcceso(3)
    Barher.Buttons("Cargos").Enabled = TieneAcceso(5)
    Barher.Buttons("Status").Enabled = TieneAcceso(7)
End Sub

Private Sub MDIForm_Activate()
    CentrarLogin
End Sub

Private Sub MDIForm_Resize()
    AjustarProgreso
    CentrarLogin
End Sub

Private Sub Salir_Click()
    MdiConect.CerrarAplicacion
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not CerrandoAplicacion Then
        If Not MdiConect.PuedeCerrarFormularios() Then
            Cancel = True
            Exit Sub
        End If
    End If
    CerrandoAplicacion = True
    VarUsuario = vbNullString
    EjecuteDatrav = False
End Sub

Private Sub AyuAce_Click()
    Dim Informacion As String
    Informacion = "Clever SoftPro X6 - LSLA" & vbCrLf & "Versión interna 6.1 · Mayo de 2011" & vbCrLf & vbCrLf
    Informacion = Informacion & "Sistema de gestión para un courier local, desarrollado para Logistic & Solutions Latin American (LSLA), Lima, Perú, al inicio de sus operaciones." & vbCrLf & vbCrLf
    Informacion = Informacion & "Organiza los pedidos de los clientes y el recorrido de sus documentos: registro de servicios y guías, importación desde Excel, etiquetas, agrupación de servicios, reparto por mensajero, cargos y consolidados." & vbCrLf & vbCrLf
    Informacion = Informacion & "Permite confirmar entregas e incidencias, consultar historiales y Delivery, generar informes y administrar usuarios, perfiles y permisos." & vbCrLf & vbCrLf
    Informacion = Informacion & "Tecnología: Visual Basic 6.0, SQL Server, ADO, controles OCX, Data Report y automatización de Microsoft Excel." & vbCrLf & vbCrLf
    Informacion = Informacion & "Diseñado y desarrollado por Juan Pablo Aguilar Campos." & vbCrLf & "jaguilar.rationalsys@outlook.com"
    MsgBox Informacion, vbInformation, "Acerca de Clever SoftPro X6"
End Sub

Private Sub ayuCon_Click()
    CmdAyuda
End Sub

Private Sub CmdAyuda()
    MsgBox "Inicie sesion con su usuario de Clever. Use Archivo > Salir para cerrar el sistema.", vbInformation, "Clever SoftPro"
End Sub
Public Sub ProgramarCierre()
    TmrCerrar.Enabled = True
End Sub

Private Sub TmrCerrar_Timer()
    TmrCerrar.Enabled = False
    MdiConect.CerrarAplicacion
End Sub

Private Sub Mantenimiento_Click()
    If TieneAcceso(1) Then FrmManten.Show
End Sub

Private Sub BarHer_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Clientes": Mantenimiento_Click
        Case "Registros": OpeVis_Click
        Case "Etiquetas": Registros_Click
        Case "Cargos": Reparto_Click
        Case "Status": RegStatus_Click
    End Select
End Sub

Private Sub BusReg_Click()
    frmSerch.Show vbModal, FrmMenuFlot
End Sub

Private Sub busEli_Click()
    On Error GoTo ErrInt
    FrmMenuFlot.FiltrarRegistros vbNullString
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub
Private Sub OpeVis_Click()
    If TieneAcceso(2) Then FrmVisDocument.Show
End Sub

Private Sub Registros_Click()
    If TieneAcceso(3) Then FrmEtiquetas.Show
End Sub

Private Sub OpeGrup_Click()
    If TieneAcceso(4) Then FrmgestServ.Show
End Sub

Private Sub Reparto_Click()
    If TieneAcceso(5) Then FrmReparto.Show
End Sub

Private Sub Consol_Click()
    If TieneAcceso(6) Then Frmconsolidado.Show
End Sub

Private Sub RegStatus_Click()
    If TieneAcceso(7) Then FrmConfirm.Show
End Sub

Private Sub Cntdeli_Click()
    If TieneAcceso(8) Then FrmDelivery.Show
End Sub

Private Sub ForStatus_Click()
    If TieneAcceso(9) Then Inform_Status.Show
End Sub

Private Sub ConfVar_Click()
    If TieneAcceso(10) Then FrmConfig.Show
End Sub

Private Sub confActu_Click()
    If TieneAcceso(11) Then FrmActualizaciones.Show vbModal, Me
End Sub

Private Sub AdministrarUsuarios_Click()
    On Error GoTo ErrInt
    ExigirAdministrador
    FrmUsuarios.Show
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub
