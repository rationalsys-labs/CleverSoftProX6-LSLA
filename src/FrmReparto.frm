VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FrmReparto 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Reparto y Ranches"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10800
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10800
   Begin VB.CommandButton CmdCancelar
      Caption         =   "Nueva Lista"
      Height          =   375
      Left            =   6885
      TabIndex        =   11
      Top             =   6750
      Width           =   1620
   End
   Begin VB.CommandButton CmdUltimoCargo
      Caption         =   "Ultimo Cargo"
      Height          =   375
      Left            =   5445
      TabIndex        =   12
      Top             =   6750
      Width           =   1305
   End
   Begin VB.Frame Frame1 
      Height          =   555
      Left            =   135
      TabIndex        =   6
      Top             =   6615
      Width           =   5100
      Begin VB.CommandButton CmdRegCarg 
         Caption         =   "Regenerar"
         Enabled         =   0   'False
         Height          =   285
         Left            =   3735
         TabIndex        =   9
         Top             =   180
         Width           =   1275
      End
      Begin VB.TextBox TxtRegCarg 
         Alignment       =   2  'Center
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         Height          =   285
         Left            =   2520
         TabIndex        =   8
         Top             =   180
         Width           =   1140
      End
      Begin VB.CheckBox ChkReg 
         Caption         =   "Regenerar Cargo de Reparto"
         Height          =   195
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   2355
      End
   End
   Begin VB.ComboBox CmbMensajeros 
      Height          =   315
      Left            =   7020
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   90
      Width           =   3660
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   3
      Top             =   7215
      Width           =   10800
      _ExtentX        =   19050
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
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   540
      Left            =   4905
      ScaleHeight     =   510
      ScaleWidth      =   1620
      TabIndex        =   4
      Top             =   7155
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.CommandButton CmdRegRep 
      Caption         =   "Registrar Reparto"
      Height          =   375
      Left            =   8640
      TabIndex        =   2
      Top             =   6750
      Width           =   2040
   End
   Begin VB.TextBox txtregistro 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   90
      Width           =   2085
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6180
      Left            =   135
      TabIndex        =   0
      Top             =   450
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   10901
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin VB.Label LblRanche 
      BackColor       =   &H00808080&
      Height          =   285
      Left            =   5445
      TabIndex        =   10
      Top             =   6750
      Visible         =   0   'False
      Width           =   1185
   End
End
Attribute VB_Name = "FrmReparto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Guardado As Boolean
Private Cambiando As Boolean
Private Guardando As Boolean
Private UltimoCargo As String
Private FechaCargo As Date

Private Sub Form_Load()
    Dim Panel As MSComctlLib.Panel
    On Error GoTo ErrInt
    Me.Icon = MDIPri.ICO.ListImages(25).Picture
    For Each Panel In StatusBar1.Panels: Panel.MinWidth = 15: Next
    StatusBar1.Panels(1).Width = 2550: StatusBar1.Panels(2).AutoSize = sbrSpring: StatusBar1.Panels(3).Width = 2850
    PrepararCuadriculaReparto Grid, False
    CargarLista CmbMensajeros, "SELECT CodMen,NomMen FROM dbo.Mdbmensajeros WHERE Active=1 ORDER BY NomMen"
    txtregistro.MaxLength = 18: TxtRegCarg.MaxLength = 18
    txtregistro.ToolTipText = "Codigo de guia: Enter para agregar. Suprimir retira la fila pendiente."
    ActualizarEstado
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub ActualizarEstado()
    Dim Regenerando As Boolean
    Regenerando = (ChkReg.Value = 1)
    CmbMensajeros.Enabled = (Not Regenerando And Not Guardado And Not Guardando)
    txtregistro.Enabled = CmbMensajeros.Enabled
    ChkReg.Enabled = Not Guardando: TxtRegCarg.Enabled = (Regenerando And Not Guardando)
    If Regenerando Then TxtRegCarg.BackColor = vbWindowBackground Else TxtRegCarg.BackColor = &H808080
    CmdRegRep.Enabled = (Not Regenerando And Not Guardando And Grid.Rows > 1 And (Guardado Or CmbMensajeros.ListIndex >= 0))
    CmdRegCarg.Enabled = (Regenerando And Guardado And Grid.Rows > 1 And Not Guardando)
    CmdCancelar.Enabled = Not Guardando
    CmdUltimoCargo.Enabled = (Len(UltimoCargo) > 0 And Not Guardando)
    If Guardado Then CmdRegRep.Caption = "Reimprimir Cargo" Else CmdRegRep.Caption = "Registrar Reparto"
    If Regenerando Then CmdCancelar.Caption = "Volver" Else CmdCancelar.Caption = "Nueva Lista"
    If Guardando Then
        StatusBar1.Panels(1).Text = "Guardando reparto"
    ElseIf Regenerando Then
        If Guardado Then StatusBar1.Panels(1).Text = "Cargo consultado: " & LblRanche.Caption Else StatusBar1.Panels(1).Text = "Ingrese cargo y Enter"
    ElseIf Guardado Then
        StatusBar1.Panels(1).Text = "Cargo guardado: " & LblRanche.Caption
    Else
        StatusBar1.Panels(1).Text = "Pendientes: " & CStr(Grid.Rows - 1)
    End If
    If Regenerando Then
        StatusBar1.Panels(2).Text = "Regenerar imprime; Volver sale de la consulta"
    ElseIf CmbMensajeros.ListIndex >= 0 Then
        StatusBar1.Panels(2).Text = CmbMensajeros.Text & " | Enter: agregar; Supr: retirar"
    Else
        StatusBar1.Panels(2).Text = "Seleccione mensajero antes de registrar"
    End If
    If Len(UltimoCargo) > 0 Then StatusBar1.Panels(3).Text = "Ultimo cargo: " & UltimoCargo Else StatusBar1.Panels(3).Text = "Guias: " & CStr(Grid.Rows - 1)
End Sub

Public Sub AgregarDocumento(ByVal Codigo As String)
    If Guardando Or ChkReg.Value = 1 Then Err.Raise 5, , "Inicie una lista de reparto para agregar guias."
    If Guardado Then Err.Raise 5, , "El cargo ya esta guardado. Use Nueva Lista o reimprima el cargo."
    AgregarDocumentoReparto Grid, Codigo, False, CmbMensajeros.Text
    txtregistro.Text = ""
    ActualizarEstado
End Sub

Public Sub GuardarCargo()
    Dim Fila As Long, Numero As Long, Descripcion As String
    If Guardado Or Guardando Then Exit Sub
    If CmbMensajeros.ListIndex < 0 Then Err.Raise 5, , "Seleccione un mensajero."
    If Len(Trim$(txtregistro.Text)) > 0 Then Err.Raise 5, , "Hay un codigo sin agregar. Presione Enter o borrelo antes de registrar."
    On Error GoTo ErrInt
    Guardando = True: ActualizarEstado
    LblRanche.Caption = RegistrarReparto(Grid, CmbMensajeros.ItemData(CmbMensajeros.ListIndex), 1, FechaCargo)
    Guardado = True: UltimoCargo = LblRanche.Caption
    For Fila = 1 To Grid.Rows - 1: Grid.TextMatrix(Fila, 7) = CmbMensajeros.Text: Next
    Guardando = False: ActualizarEstado
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    Guardando = False: ActualizarEstado
    Err.Raise Numero, "GuardarCargo", Descripcion
End Sub

Public Sub CargarCargo(ByVal Codigo As String)
    Dim Datos As ADODB.Recordset, Fila As Long, Columna As Integer
    Dim Numero As Long, Descripcion As String
    On Error GoTo ErrInt
    Grid.Rows = 1: LblRanche.Caption = "": Guardado = False: FechaCargo = 0
    ActualizarEstado
    If Not CodigoValido(Codigo) Then Err.Raise 5, , "Ingrese un numero de cargo valido."
    Set Datos = ConsultarDatos("SELECT Codigo,Producto,Empresa,Direccion,Distrito,Contacto,Fecha,Informe,Cliente FROM dbo.ViewRegenRanche WHERE Ranche=? ORDER BY Codigo", Codigo, True)
    If Datos.EOF Then Err.Raise 5, , "No se encontraron registros activos de ese cargo."
    Do Until Datos.EOF
        If FechaCargo = 0 Then FechaCargo = CDate(Datos.Fields("Fecha").Value)
        If CDate(Datos.Fields("Fecha").Value) < FechaCargo Then FechaCargo = CDate(Datos.Fields("Fecha").Value)
        Grid.Rows = Grid.Rows + 1: Fila = Grid.Rows - 1
        For Columna = 0 To 8: Grid.TextMatrix(Fila, Columna) = TextoDato(Datos.Fields(Columna).Value): Next
        Datos.MoveNext
    Loop
    Datos.Close
    LblRanche.Caption = CStr(CDec(Codigo)): Guardado = True
    ActualizarEstado
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    Grid.Rows = 1: LblRanche.Caption = "": Guardado = False
    On Error GoTo 0
    ActualizarEstado
    Err.Raise Numero, "CargarCargo", Descripcion
End Sub

Private Sub CmdRegRep_Click()
    On Error GoTo ErrInt
    If Not CmdRegRep.Enabled Or Guardando Then Exit Sub
    If Len(Trim$(txtregistro.Text)) > 0 Then Err.Raise 5, , "Hay un codigo sin agregar. Presione Enter o borrelo antes de continuar."
    If Guardado Then
        If MsgBox("Se reimprimira el cargo " & LblRanche.Caption & ", sin registrar otro reparto. Desea continuar?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") <> vbYes Then Exit Sub
    Else
        If MsgBox("Se registrara e imprimira el reparto de " & CStr(Grid.Rows - 1) & " guias para " & CmbMensajeros.Text & ". Desea continuar?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") <> vbYes Then Exit Sub
        GuardarCargo
    End If
    ImprimirCargo Grid, LblRanche.Caption, FechaCargo, False
    LimpiarLista
    Exit Sub
ErrInt:
    If Guardado Then
        MsgBox "El cargo " & LblRanche.Caption & " permanece guardado. Puede reintentar la impresion sin volver a registrarlo." & vbCrLf & Err.Description, vbExclamation, "Clever SoftPro"
    Else
        MsgBox Err.Description, vbExclamation, "Clever SoftPro"
    End If
End Sub

Private Sub CmdRegCarg_Click()
    On Error GoTo ErrInt
    If Not CmdRegCarg.Enabled Then Exit Sub
    If MsgBox("Se imprimira el cargo " & LblRanche.Caption & ". Desea continuar?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") <> vbYes Then Exit Sub
    ImprimirCargo Grid, LblRanche.Caption, FechaCargo, True
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub LimpiarLista()
    Grid.Rows = 1: Guardado = False: LblRanche.Caption = "": FechaCargo = 0
    txtregistro.Text = ""
    ActualizarEstado
End Sub

Private Sub CmdCancelar_Click()
    If Not PuedeCerrar() Then Exit Sub
    Cambiando = True: ChkReg.Value = 0: Cambiando = False
    TxtRegCarg.Text = "": LimpiarLista
End Sub

Private Sub CmdUltimoCargo_Click()
    On Error GoTo ErrInt
    If Len(UltimoCargo) = 0 Or Not PuedeCerrar() Then Exit Sub
    Cambiando = True: ChkReg.Value = 1: Cambiando = False
    LimpiarLista
    TxtRegCarg.Text = UltimoCargo
    CargarCargo UltimoCargo
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub ChkReg_Click()
    If Cambiando Then Exit Sub
    If Not PuedeCerrar() Then
        Cambiando = True: ChkReg.Value = 1 - ChkReg.Value: Cambiando = False
        Exit Sub
    End If
    LimpiarLista: TxtRegCarg.Text = ""
    ActualizarEstado
End Sub

Private Sub txtregistro_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrInt
    Select Case KeyAscii
        Case 13: KeyAscii = 0: AgregarDocumento Trim$(txtregistro.Text)
        Case 48 To 57, 8, 22
        Case Else: KeyAscii = 0
    End Select
    Exit Sub
ErrInt:
    If Err.Number = vbObjectError + 401 Then
        If MsgBox(Err.Description & vbCrLf & "Desea consultar la liberacion de la guia?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") = vbYes Then
            Load FrmReactivar: FrmReactivar.PrepararLiberacion txtregistro.Text: FrmReactivar.Show vbModal, MDIPri
        End If
    Else
        MsgBox Err.Description, vbExclamation, "Clever SoftPro"
    End If
End Sub

Private Sub TxtRegCarg_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrInt
    Select Case KeyAscii
        Case 13: KeyAscii = 0: CargarCargo Trim$(TxtRegCarg.Text)
        Case 48 To 57, 8, 22
        Case Else: KeyAscii = 0
    End Select
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmbMensajeros_Click()
    Dim Fila As Long
    If Guardado Or Guardando Then Exit Sub
    For Fila = 1 To Grid.Rows - 1: Grid.TextMatrix(Fila, 7) = CmbMensajeros.Text: Next
    ActualizarEstado
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Or Guardado Or Guardando Or Grid.Row < 1 Or Grid.Rows < 2 Then Exit Sub
    KeyCode = 0
    If MsgBox("Desea quitar la guia " & Grid.TextMatrix(Grid.Row, 0) & " de esta lista pendiente? No se borrara el documento.", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") = vbYes Then QuitarDocumentoReparto Grid
    ActualizarEstado
End Sub

Public Function PuedeCerrar() As Boolean
    If Guardando Then Exit Function
    PuedeCerrar = True
    If Grid.Rows < 2 And Len(txtregistro.Text) = 0 Then Exit Function
    If ChkReg.Value = 1 And Guardado Then Exit Function
    If Guardado Then
        PuedeCerrar = (MsgBox("El cargo " & LblRanche.Caption & " esta guardado y puede regenerarse. Desea cerrar esta lista?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") = vbYes)
    Else
        PuedeCerrar = (MsgBox("Hay guias o un codigo sin registrar. Desea descartar la lista pendiente?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") = vbYes)
    End If
    If PuedeCerrar Then LimpiarLista
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then KeyCode = 0: CmdCancelar_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not CerrandoAplicacion Then Cancel = Not PuedeCerrar()
End Sub

Private Sub TxtRegCarg_Change()
    If ChkReg.Value <> 1 Then Exit Sub
    Grid.Rows = 1: LblRanche.Caption = "": Guardado = False: FechaCargo = 0
    ActualizarEstado
End Sub

Private Sub Form_Activate()
    VerificarAccesoFormulario Me, 5
End Sub
