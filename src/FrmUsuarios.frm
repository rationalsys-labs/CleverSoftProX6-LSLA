VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmUsuarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Administracion de Usuarios"
   ClientHeight    =   4065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10215
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4065
   ScaleWidth      =   10215
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   0
      Top             =   3735
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "9/11/2026"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Cmd_Inf 
      Caption         =   "Listar Usuarios"
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   2535
   End
   Begin VB.CommandButton Cmd_Mod 
      Caption         =   "Modificar Usuarios"
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton Cmd_New 
      Caption         =   "Crear Usuario"
      Height          =   375
      Left            =   4560
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8280
      TabIndex        =   40
      Top             =   120
      Width           =   1800
   End
   Begin VB.Frame Fra_New 
      Height          =   3135
      Left            =   120
      TabIndex        =   23
      Top             =   480
      Visible         =   0   'False
      Width           =   9975
      Begin VB.Frame Frame1 
         Height          =   975
         Left            =   240
         TabIndex        =   30
         Top             =   1560
         Width           =   7215
         Begin VB.TextBox NewPasCon 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   3840
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   32
            Top             =   480
            Width           =   3135
         End
         Begin VB.TextBox NewPas 
            Appearance      =   0  'Flat
            Height          =   285
            IMEMode         =   3  'DISABLE
            Left            =   240
            MaxLength       =   15
            PasswordChar    =   "*"
            TabIndex        =   31
            Top             =   480
            Width           =   3255
         End
         Begin VB.Label Label12 
            Caption         =   "Confirmar Contraseña"
            Height          =   255
            Left            =   3840
            TabIndex        =   34
            Top             =   240
            Width           =   1575
         End
         Begin VB.Label Label11 
            Caption         =   "Contraseña"
            Height          =   255
            Left            =   240
            TabIndex        =   33
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.TextBox NewMail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   40
         TabIndex        =   29
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox NewUsu 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   5
         TabIndex        =   28
         Top             =   720
         Width           =   2895
      End
      Begin VB.TextBox NewApe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         MaxLength       =   20
         TabIndex        =   27
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox NewNom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         MaxLength       =   20
         TabIndex        =   26
         Top             =   360
         Width           =   2895
      End
      Begin VB.ComboBox NewAccess 
         Height          =   315
         ItemData        =   "FrmUsuarios.frx":0000
         Left            =   5040
         List            =   "FrmUsuarios.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   1080
         Width           =   2415
      End
      Begin VB.TextBox NewSit 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   24
         Text            =   "Activo"
         Top             =   720
         Width           =   2415
      End
      Begin VB.CommandButton CmdNew 
         Caption         =   "Crear Usuario"
         Height          =   375
         Left            =   7680
         TabIndex        =   41
         Top             =   360
         Width           =   2055
      End
      Begin VB.CommandButton CmdLim 
         Caption         =   "Limpiar Datos"
         Height          =   375
         Left            =   7680
         TabIndex        =   42
         Top             =   840
         Width           =   2055
      End
      Begin VB.Label Label7 
         Caption         =   "Email:"
         Height          =   255
         Left            =   240
         TabIndex        =   43
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   240
         TabIndex        =   39
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Apellido:"
         Height          =   255
         Left            =   4320
         TabIndex        =   38
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label10 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   240
         TabIndex        =   37
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Acceso:"
         Height          =   255
         Left            =   4320
         TabIndex        =   36
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label13 
         Caption         =   "Situacion:"
         Height          =   255
         Left            =   4320
         TabIndex        =   35
         Top             =   720
         Width           =   615
      End
   End
   Begin VB.Frame Fra_Mod 
      Height          =   3135
      Left            =   120
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   9975
      Begin VB.TextBox TxtNom 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Top             =   360
         Width           =   3015
      End
      Begin VB.TextBox TxtApe 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox TxtUsu 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Top             =   1080
         Width           =   1575
      End
      Begin VB.TextBox TxtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   1440
         Width           =   3015
      End
      Begin VB.ComboBox CmbSit 
         Height          =   315
         ItemData        =   "FrmUsuarios.frx":0004
         Left            =   1080
         List            =   "FrmUsuarios.frx":000E
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1800
         Width           =   1575
      End
      Begin VB.ComboBox CmbPerfil 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   2250
         Width           =   2400
      End
      Begin VB.ListBox LstUsuario 
         Appearance      =   0  'Flat
         Height          =   1785
         ItemData        =   "FrmUsuarios.frx":0024
         Left            =   4440
         List            =   "FrmUsuarios.frx":0026
         TabIndex        =   9
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton CmdGuaMod 
         Caption         =   "Guardar Cambios"
         Height          =   375
         Left            =   7080
         TabIndex        =   20
         Top             =   360
         Width           =   2655
      End
      Begin VB.CommandButton CmdCaducar 
         Caption         =   "Cambiar Clave"
         Height          =   375
         Left            =   7080
         TabIndex        =   21
         Top             =   840
         Width           =   2655
      End
      Begin VB.CommandButton Boton5 
         Caption         =   "Dar de Baja"
         Height          =   375
         Left            =   7080
         TabIndex        =   22
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label LblPerfil 
         Caption         =   "Perfil:"
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   2250
         Width           =   800
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Apellido:"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Usuario:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Email:"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1440
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Situación:"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   735
      End
   End
   Begin VB.Frame Fra_inf 
      Height          =   3135
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   9975
      Begin MSFlexGridLib.MSFlexGrid GridUsu 
         Height          =   2775
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   6975
         _ExtentX        =   12303
         _ExtentY        =   4895
         _Version        =   393216
         FixedCols       =   0
         Appearance      =   0
      End
      Begin VB.CommandButton CmdPermisos 
         Caption         =   "Permisos de Usuario"
         Height          =   375
         Left            =   7200
         TabIndex        =   3
         Top             =   240
         Width           =   2655
      End
      Begin VB.CommandButton Boton1 
         Caption         =   "Actualizar Lista"
         Height          =   375
         Left            =   7200
         TabIndex        =   4
         Top             =   720
         Width           =   2655
      End
   End
End
Attribute VB_Name = "FrmUsuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Codigo As Long
Private Cambio As Boolean
Private Cargando As Boolean

Private Sub Form_Load()
    On Error GoTo ErrorCarga
    ExigirAdministrador
    Me.Icon = MDIPri.ICO.ListImages(25).Picture
    NewNom.MaxLength = 50: NewApe.MaxLength = 50: NewMail.MaxLength = 150
    TxtNom.MaxLength = 50: TxtApe.MaxLength = 50: TxtEmail.MaxLength = 150
    NewMail.ToolTipText = "Correo opcional, hasta 150 caracteres.": TxtEmail.ToolTipText = NewMail.ToolTipText
    NewPas.MaxLength = 150: NewPasCon.MaxLength = 150
    NewAccess.Clear: NewAccess.AddItem "Administrador": NewAccess.AddItem "Operador"
    CmbPerfil.AddItem "Administrador": CmbPerfil.AddItem "Operador"
    CmbSit.Clear: CmbSit.AddItem "Activo": CmbSit.AddItem "Inactivo"
    CargarUsuariosLista
    Exit Sub
ErrorCarga:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Public Sub CargarUsuariosLista()
    Dim Datos As ADODB.Recordset, Fila As Long, Columna As Integer, Campos As Variant
    On Error GoTo ErrorCarga
    Cargando = True: Codigo = 0
    GridUsu.Rows = 1: GridUsu.Cols = 8: LstUsuario.Clear
    Campos = Array("Codigo", "Nombres", "Apellidos", "Usuario", "Perfil", "Situacion", "Clave vigente", "Correo")
    For Columna = 0 To 7
        GridUsu.TextMatrix(0, Columna) = Campos(Columna): GridUsu.ColWidth(Columna) = 1600
    Next
    GridUsu.ColWidth(0) = 0: GridUsu.ColWidth(3) = 800
    Set Datos = ListarUsuarios()
    Do Until Datos.EOF
        GridUsu.Rows = GridUsu.Rows + 1: Fila = GridUsu.Rows - 1
        GridUsu.TextMatrix(Fila, 0) = CStr(Datos!Item)
        GridUsu.TextMatrix(Fila, 1) = TextoDato(Datos!NomUsu): GridUsu.TextMatrix(Fila, 2) = TextoDato(Datos!ApeUsu)
        GridUsu.TextMatrix(Fila, 3) = CStr(Datos!Usuari)
        GridUsu.TextMatrix(Fila, 4) = IIf(CLng(Datos.Fields("Access").value) = 0, "Administrador", "Operador")
        GridUsu.TextMatrix(Fila, 5) = IIf(CBool(Datos!SitUsu), "Activo", "Inactivo")
        GridUsu.TextMatrix(Fila, 6) = IIf(CBool(Datos!ActPas), "Si", "No")
        GridUsu.TextMatrix(Fila, 7) = TextoDato(Datos!Email)
        LstUsuario.AddItem CStr(Datos!Usuari): LstUsuario.ItemData(LstUsuario.NewIndex) = CLng(Datos!Item)
        Datos.MoveNext
    Loop
    Datos.Close
    TxtNom.Text = "": TxtApe.Text = "": TxtUsu.Text = "": TxtEmail.Text = ""
    CmbSit.ListIndex = -1: CmbPerfil.ListIndex = -1
    Cargando = False: Cambio = False
    MostrarModo 0
    Exit Sub
ErrorCarga:
    Dim Descripcion As String
    Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    Cargando = False
    MsgBox Descripcion, vbExclamation, "Clever SoftPro"
End Sub

Private Sub MostrarModo(ByVal Modo As Integer)
    Fra_inf.Visible = (Modo = 0): Fra_New.Visible = (Modo = 1): Fra_Mod.Visible = (Modo = 2)
    CmdGuaMod.Enabled = (Codigo > 0): CmdCaducar.Enabled = (Codigo > 0): Boton5.Enabled = (Codigo > 0)
    CmdPermisos.Enabled = (Codigo > 0)
    ActualizarBotones
    Dim Panel As MSComctlLib.Panel
    For Each Panel In StatusBar1.Panels: Panel.MinWidth = 15: Panel.Style = sbrText: Next
    StatusBar1.Panels(1).Width = 2400: StatusBar1.Panels(3).Width = 1800
    StatusBar1.Panels(3).Text = CStr(GridUsu.Rows - 1) & " usuarios"
    StatusBar1.Panels(2).MinWidth = 15: StatusBar1.Panels(2).AutoSize = sbrSpring
    StatusBar1.Panels(1).Text = Choose(Modo + 1, "Consulta de usuarios", "Creacion de usuario", "Modificacion de usuario")
    If Codigo > 0 Then StatusBar1.Panels(2).Text = "Usuario seleccionado: " & TxtUsu.Text Else StatusBar1.Panels(2).Text = "Seleccione una cuenta. La baja conserva su historial."
End Sub

Public Function PuedeCerrar() As Boolean
    PuedeCerrar = True
    If Cambio Then PuedeCerrar = (MsgBox("Hay datos de usuario sin guardar. Desea descartarlos?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") = vbYes)
End Function

Private Sub Cmd_Inf_Click()
    If Not PuedeCerrar() Then Exit Sub
    Cambio = False
    If Codigo > 0 Then SeleccionarUsuario Codigo
    MostrarModo 0
End Sub

Private Sub Cmd_Mod_Click()
    If Not PuedeCerrar() Then Exit Sub
    Cambio = False
    If Codigo > 0 Then SeleccionarUsuario Codigo
    MostrarModo 2
End Sub

Private Sub Cmd_New_Click()
    If Not PuedeCerrar() Then Exit Sub
    Cargando = True
    NewNom.Text = "": NewApe.Text = "": NewUsu.Text = "": NewMail.Text = "": NewPas.Text = "": NewPasCon.Text = ""
    NewAccess.ListIndex = 1: NewSit.Text = "Activo"
    Cargando = False: Cambio = False: MostrarModo 1
End Sub

Private Sub CmdCancelar_Click()
    If PuedeCerrar() Then CargarUsuariosLista
End Sub

Private Sub CmdLim_Click()
    Cmd_New_Click
End Sub

Private Sub CmdNew_Click()
    On Error GoTo ErrorGuardar
    If NewAccess.ListIndex < 0 Then Err.Raise 5, , "Seleccione el perfil."
    Codigo = GuardarUsuario(0, 1, NewNom.Text, NewApe.Text, NewUsu.Text, NewPas.Text, NewPasCon.Text, NewAccess.ListIndex, True, NewMail.Text)
    Cambio = False: NewPas.Text = "": NewPasCon.Text = ""
    CargarUsuariosLista
    MsgBox "Usuario creado. Los operadores requieren que se les asignen permisos por modulo.", vbInformation, "Clever SoftPro"
    Exit Sub
ErrorGuardar:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Public Sub SeleccionarUsuario(ByVal Clave As Long)
    Dim Fila As Long
    If Not PuedeCerrar() Then
        RestaurarSeleccion
        Exit Sub
    End If
    For Fila = 1 To GridUsu.Rows - 1
        If CLng(GridUsu.TextMatrix(Fila, 0)) = Clave Then
            Cargando = True: Codigo = Clave
            TxtNom.Text = GridUsu.TextMatrix(Fila, 1): TxtApe.Text = GridUsu.TextMatrix(Fila, 2)
            TxtUsu.Text = GridUsu.TextMatrix(Fila, 3): TxtEmail.Text = GridUsu.TextMatrix(Fila, 7)
            CmbPerfil.ListIndex = IIf(GridUsu.TextMatrix(Fila, 4) = "Administrador", 0, 1)
            CmbSit.ListIndex = IIf(GridUsu.TextMatrix(Fila, 5) = "Activo", 0, 1)
            RestaurarSeleccion
            Cargando = False: Cambio = False
            If Fra_Mod.Visible Then MostrarModo 2 Else MostrarModo 0
            Exit Sub
        End If
    Next
End Sub

Private Sub RestaurarSeleccion()
    Dim Fila As Long
    Cargando = True
    LstUsuario.ListIndex = -1
    GridUsu.Row = 0: GridUsu.RowSel = 0
    For Fila = 0 To LstUsuario.ListCount - 1
        If LstUsuario.ItemData(Fila) = Codigo Then LstUsuario.ListIndex = Fila: Exit For
    Next
    For Fila = 1 To GridUsu.Rows - 1
        If CLng(GridUsu.TextMatrix(Fila, 0)) = Codigo Then
            GridUsu.Row = Fila: GridUsu.RowSel = Fila
            Exit For
        End If
    Next
    Cargando = False
End Sub

Private Sub LstUsuario_Click()
    If Not Cargando And LstUsuario.ListIndex >= 0 Then SeleccionarUsuario LstUsuario.ItemData(LstUsuario.ListIndex)
End Sub

Private Sub GridUsu_Click()
    If Not Cargando And GridUsu.Row > 0 And GridUsu.Rows > 1 Then SeleccionarUsuario CLng(GridUsu.TextMatrix(GridUsu.Row, 0))
End Sub

Private Sub CmdGuaMod_Click()
    On Error GoTo ErrorGuardar
    If Codigo = 0 Or CmbPerfil.ListIndex < 0 Or CmbSit.ListIndex < 0 Then Err.Raise 5, , "Seleccione el usuario, perfil y situacion."
    GuardarUsuario Codigo, 2, TxtNom.Text, TxtApe.Text, , , , CmbPerfil.ListIndex, (CmbSit.ListIndex = 0), TxtEmail.Text
    Cambio = False: CargarUsuariosLista
    Exit Sub
ErrorGuardar:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub Boton5_Click()
    On Error GoTo ErrorGuardar
    If Codigo = 0 Then Err.Raise 5, , "Seleccione el usuario que desea dar de baja."
    If MsgBox("Desea dar de baja al usuario " & TxtUsu.Text & " (codigo " & CStr(Codigo) & "), " & TxtNom.Text & " " & TxtApe.Text & "? Se conservara su historial y se impedira un nuevo inicio de sesion.", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") <> vbYes Then Exit Sub
    GuardarUsuario Codigo, 3
    Cambio = False: CargarUsuariosLista
    Exit Sub
ErrorGuardar:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdCaducar_Click()
    On Error GoTo ErrorGuardar
    If Codigo = 0 Then Err.Raise 5, , "Seleccione el usuario."
    Load FrmClaveUsuario: FrmClaveUsuario.Preparar Codigo, TxtUsu.Text
    FrmClaveUsuario.Show vbModal, MDIPri
    Exit Sub
ErrorGuardar:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdPermisos_Click()
    On Error GoTo ErrorGuardar
    If Codigo = 0 Then Err.Raise 5, , "Seleccione el usuario."
    Load FrmPermisos: FrmPermisos.Preparar Codigo, TxtUsu.Text, (CmbPerfil.ListIndex = 0)
    FrmPermisos.Show vbModal, MDIPri
    Exit Sub
ErrorGuardar:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub Boton1_Click()
    If PuedeCerrar() Then CargarUsuariosLista
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not CerrandoAplicacion Then Cancel = Not PuedeCerrar()
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then KeyCode = 0: CmdCancelar_Click
End Sub

Private Sub NewNom_Change()
    If Not Cargando Then Cambio = True: ActualizarBotones
End Sub

Private Sub NewApe_Change()
    If Not Cargando Then Cambio = True: ActualizarBotones
End Sub

Private Sub NewUsu_Change()
    If Not Cargando Then Cambio = True: ActualizarBotones
End Sub

Private Sub NewMail_Change()
    If Not Cargando Then Cambio = True: ActualizarBotones
End Sub

Private Sub NewPas_Change()
    If Not Cargando Then Cambio = True: ActualizarBotones
End Sub

Private Sub NewPasCon_Change()
    If Not Cargando Then Cambio = True: ActualizarBotones
End Sub

Private Sub TxtNom_Change()
    If Not Cargando Then Cambio = True: ActualizarBotones
End Sub

Private Sub TxtApe_Change()
    If Not Cargando Then Cambio = True: ActualizarBotones
End Sub

Private Sub TxtEmail_Change()
    If Not Cargando Then Cambio = True: ActualizarBotones
End Sub

Private Sub NewAccess_Click()
    If Not Cargando Then Cambio = True: ActualizarBotones
End Sub

Private Sub CmbSit_Click()
    If Not Cargando Then Cambio = True: ActualizarBotones
End Sub

Private Sub CmbPerfil_Click()
    If Not Cargando Then Cambio = True: ActualizarBotones
End Sub

Private Sub Form_Activate()
    VerificarAccesoFormulario Me, 0
End Sub

Private Sub ActualizarBotones()
    If Cargando Then Exit Sub
    CmdNew.Enabled = (Len(Trim$(NewNom.Text)) > 0 And Len(Trim$(NewApe.Text)) > 0 And Len(Trim$(NewUsu.Text)) > 0 And NewAccess.ListIndex >= 0 And Len(NewPas.Text) >= 6 And NewPas.Text = NewPasCon.Text)
    CmdGuaMod.Enabled = (Codigo > 0 And Cambio And Len(Trim$(TxtNom.Text)) > 0 And Len(Trim$(TxtApe.Text)) > 0 And CmbSit.ListIndex >= 0 And CmbPerfil.ListIndex >= 0)
End Sub
