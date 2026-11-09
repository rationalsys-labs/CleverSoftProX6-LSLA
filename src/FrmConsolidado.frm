VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FrmConsolidado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Consolidados"
   ClientHeight    =   7470
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13470
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7470
   ScaleWidth      =   13470
   Begin VB.CommandButton CmdCancelar
      Caption         =   "Nueva Lista"
      Height          =   375
      Left            =   9360
      TabIndex        =   5
      Top             =   6750
      Width           =   1725
   End
   Begin VB.CommandButton CmdUltimoInforme
      Caption         =   "Ultimo Informe"
      Height          =   375
      Left            =   135
      TabIndex        =   6
      Top             =   6750
      Width           =   2500
   End
   Begin VB.CommandButton CmdRegRep 
      Caption         =   "Registrar Reparto"
      Height          =   375
      Left            =   11250
      TabIndex        =   3
      Top             =   6750
      Width           =   2040
   End
   Begin VB.TextBox txtregistro 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   90
      TabIndex        =   1
      Top             =   135
      Width           =   2085
   End
   Begin VB.ComboBox CmbMensajeros 
      Height          =   315
      Left            =   9585
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   135
      Width           =   3660
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   6180
      Left            =   135
      TabIndex        =   2
      Top             =   495
      Width           =   13155
      _ExtentX        =   23204
      _ExtentY        =   10901
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   4
      Top             =   7230
      Width           =   13470
      _ExtentX        =   23760
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
End
Attribute VB_Name = "FrmConsolidado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Guardado As Boolean
Private Guardando As Boolean
Private FechaRegistro As Date
Private UltimosDatos As ADODB.Recordset
Private UltimaFecha As Date

Private Sub Form_Load()
    Dim Panel As MSComctlLib.Panel
    On Error GoTo ErrInt
    Me.Icon = MDIPri.ICO.ListImages(25).Picture
    For Each Panel In StatusBar1.Panels: Panel.MinWidth = 15: Next
    StatusBar1.Panels(1).Width = 2700: StatusBar1.Panels(2).AutoSize = sbrSpring: StatusBar1.Panels(3).Width = 2400
    PrepararCuadriculaReparto Grid, True
    CargarLista CmbMensajeros, "SELECT CodMen,NomMen FROM dbo.Mdbmensajeros WHERE Active=1 ORDER BY NomMen"
    txtregistro.MaxLength = 18
    txtregistro.ToolTipText = "Codigo de guia: Enter para agregar. Suprimir retira la fila pendiente."
    ActualizarEstado
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub ActualizarEstado()
    CmbMensajeros.Enabled = (Not Guardado And Not Guardando)
    txtregistro.Enabled = (CmbMensajeros.ListIndex >= 0 And Not Guardado And Not Guardando)
    CmdRegRep.Enabled = (Grid.Rows > 1 And CmbMensajeros.ListIndex >= 0 And Not Guardando)
    CmdCancelar.Enabled = Not Guardando
    CmdUltimoInforme.Enabled = False
    If Not UltimosDatos Is Nothing Then CmdUltimoInforme.Enabled = (UltimosDatos.State = adStateOpen And Not Guardando)
    If Guardado Then
        CmdRegRep.Caption = "Generar Informe": StatusBar1.Panels(1).Text = "Consolidado guardado"
    Else
        CmdRegRep.Caption = "Registrar Reparto": StatusBar1.Panels(1).Text = "Guias pendientes: " & CStr(Grid.Rows - 1)
    End If
    If Guardando Then StatusBar1.Panels(1).Text = "Guardando consolidado"
    If CmbMensajeros.ListIndex >= 0 Then StatusBar1.Panels(2).Text = CmbMensajeros.Text Else StatusBar1.Panels(2).Text = "Seleccione un mensajero"
    StatusBar1.Panels(3).Text = "Enter: agregar | Supr: retirar"
End Sub

Public Sub AgregarDocumento(ByVal Codigo As String)
    If Guardado Or Guardando Then Err.Raise 5, , "El consolidado ya esta guardado. Genere el informe o use Nueva Lista."
    If CmbMensajeros.ListIndex < 0 Then Err.Raise 5, , "Seleccione un mensajero."
    AgregarDocumentoReparto Grid, Codigo, True, CmbMensajeros.Text
    txtregistro.Text = "": ActualizarEstado
End Sub

Public Sub GuardarConsolidado()
    Dim Resultado As String, Fila As Long, Numero As Long, Descripcion As String
    If Guardado Or Guardando Then Exit Sub
    If CmbMensajeros.ListIndex < 0 Then Err.Raise 5, , "Seleccione un mensajero."
    If Len(Trim$(txtregistro.Text)) > 0 Then Err.Raise 5, , "Hay un codigo sin agregar. Presione Enter o borrelo antes de registrar."
    On Error GoTo ErrInt
    Guardando = True: ActualizarEstado
    Resultado = RegistrarReparto(Grid, CmbMensajeros.ItemData(CmbMensajeros.ListIndex), 2, FechaRegistro)
    Guardado = True
    For Fila = 1 To Grid.Rows - 1: Grid.TextMatrix(Fila, 8) = CmbMensajeros.Text: Next
    Guardando = False: ActualizarEstado
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    Guardando = False: ActualizarEstado
    Err.Raise Numero, "GuardarConsolidado", Descripcion
End Sub

Public Sub PrepararInforme()
    Dim Datos As ADODB.Recordset
    If Not Guardado Then Err.Raise 5, , "Guarde primero el consolidado."
    RepConsol.Preparar Grid, FechaRegistro
    If Not UltimosDatos Is Nothing Then
        If UltimosDatos.State <> adStateClosed Then UltimosDatos.Close
    End If
    Set Datos = RepConsol.DataSource
    Set UltimosDatos = Datos.Clone
    UltimaFecha = FechaRegistro
    ActualizarEstado
End Sub

Public Sub PrepararUltimoInforme()
    If UltimosDatos Is Nothing Then Err.Raise 5, , "Todavia no se genero un consolidado en esta sesion."
    RepConsol.PrepararCopia UltimosDatos, UltimaFecha
End Sub

Private Sub CmdUltimoInforme_Click()
    On Error GoTo ErrInt
    If Not CmdUltimoInforme.Enabled Then Exit Sub
    PrepararUltimoInforme
    RepConsol.Show
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdRegRep_Click()
    On Error GoTo ErrInt
    If Not CmdRegRep.Enabled Or Guardando Then Exit Sub
    If Len(Trim$(txtregistro.Text)) > 0 Then Err.Raise 5, , "Hay un codigo sin agregar. Presione Enter o borrelo antes de continuar."
    If Not Guardado Then
        If MsgBox("Se registraran " & CStr(Grid.Rows - 1) & " guias para " & CmbMensajeros.Text & " y se abrira el consolidado. Desea continuar?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") <> vbYes Then Exit Sub
        GuardarConsolidado
    End If
    PrepararInforme
    RepConsol.Show
    LimpiarLista
    Exit Sub
ErrInt:
    If Guardado Then
        MsgBox "El consolidado permanece guardado. Use Generar Informe para reintentarlo sin duplicar el reparto." & vbCrLf & Err.Description, vbExclamation, "Clever SoftPro"
    Else
        MsgBox Err.Description, vbExclamation, "Clever SoftPro"
    End If
End Sub

Private Sub LimpiarLista()
    Grid.Rows = 1: Guardado = False: txtregistro.Text = ""
    ActualizarEstado
End Sub

Private Sub CmdCancelar_Click()
    If PuedeCerrar() Then LimpiarLista
End Sub

Private Sub CmbMensajeros_Click()
    Dim Fila As Long
    If Guardado Or Guardando Then Exit Sub
    For Fila = 1 To Grid.Rows - 1: Grid.TextMatrix(Fila, 8) = CmbMensajeros.Text: Next
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

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Or Guardado Or Guardando Or Grid.Row < 1 Or Grid.Rows < 2 Then Exit Sub
    KeyCode = 0
    If MsgBox("Desea quitar la guia " & Grid.TextMatrix(Grid.Row, 0) & " de la lista pendiente? No se borrara el documento.", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") = vbYes Then QuitarDocumentoReparto Grid
    ActualizarEstado
End Sub

Public Function PuedeCerrar() As Boolean
    If Guardando Then Exit Function
    PuedeCerrar = True
    If Grid.Rows < 2 And Len(txtregistro.Text) = 0 Then Exit Function
    If Guardado Then
        PuedeCerrar = (MsgBox("El consolidado esta guardado. Desea cerrar esta lista? El ultimo informe generado seguira disponible en esta ventana.", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") = vbYes)
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

Private Sub Form_Unload(Cancel As Integer)
    If Not UltimosDatos Is Nothing Then
        If UltimosDatos.State <> adStateClosed Then UltimosDatos.Close
        Set UltimosDatos = Nothing
    End If
End Sub

Private Sub Form_Activate()
    VerificarAccesoFormulario Me, 6
End Sub
