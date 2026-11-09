VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form FrmConfirm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirmacion de Registros"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13785
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   13785
   Begin MSComctlLib.TreeView ViewStatus 
      Height          =   4200
      Left            =   10350
      TabIndex        =   30
      Top             =   1755
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   7408
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.OptionButton OptPer 
      Caption         =   "Reporte Personalizado"
      Enabled         =   0   'False
      Height          =   195
      Left            =   12285
      TabIndex        =   28
      Top             =   7560
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.OptionButton OptRep 
      Caption         =   "Reporte"
      Enabled         =   0   'False
      Height          =   195
      Left            =   12285
      TabIndex        =   27
      Top             =   6255
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.CommandButton CmdConfirm 
      Caption         =   "Registrar Confirmaciones"
      Enabled         =   0   'False
      Height          =   375
      Left            =   12195
      TabIndex        =   26
      Top             =   5805
      Visible         =   0   'False
      Width           =   2760
   End
   Begin VB.Frame Frame4 
      Caption         =   "Estadistica"
      Height          =   5865
      Left            =   45
      TabIndex        =   21
      Top             =   90
      Width           =   2760
      Begin VB.Label LblContReg 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         Height          =   240
         Left            =   1800
         TabIndex        =   23
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Cantidad de Registros"
         Height          =   195
         Left            =   135
         TabIndex        =   22
         Top             =   360
         Width           =   1590
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1230
      Left            =   12195
      TabIndex        =   12
      Top             =   6255
      Visible         =   0   'False
      Width           =   2760
      Begin VB.CommandButton CmdGenRep 
         Caption         =   "Generar Reporte"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1125
         TabIndex        =   15
         Top             =   675
         Width           =   1410
      End
      Begin VB.TextBox txtServRep 
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         Height          =   285
         Left            =   945
         TabIndex        =   14
         Top             =   315
         Width           =   1590
      End
      Begin VB.Label lblServ 
         BackColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   225
         TabIndex        =   29
         Top             =   765
         Visible         =   0   'False
         Width           =   555
      End
      Begin VB.Label Label5 
         Caption         =   "Servicio:"
         Height          =   240
         Left            =   135
         TabIndex        =   13
         Top             =   315
         Width           =   600
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1545
      Left            =   12195
      TabIndex        =   16
      Top             =   7560
      Visible         =   0   'False
      Width           =   2760
      Begin VB.CommandButton CmdGenPer 
         Caption         =   "Generar Reporte"
         Enabled         =   0   'False
         Height          =   330
         Left            =   1125
         TabIndex        =   19
         Top             =   1080
         Width           =   1410
      End
      Begin VB.TextBox txtRegPer 
         BackColor       =   &H00808080&
         Enabled         =   0   'False
         Height          =   285
         Left            =   945
         TabIndex        =   18
         Top             =   720
         Width           =   1590
      End
      Begin VB.Label Label7 
         Caption         =   "Ingresar Registros para el Reporte"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   180
         TabIndex        =   20
         Top             =   270
         Width           =   2400
      End
      Begin VB.Label Label6 
         Caption         =   "Registro:"
         Height          =   240
         Left            =   180
         TabIndex        =   17
         Top             =   720
         Width           =   645
      End
   End
   Begin VB.Frame fraConf 
      Caption         =   "Registrar Confirmaciones"
      Height          =   1635
      Left            =   2880
      TabIndex        =   1
      Top             =   45
      Width           =   10770
      Begin VB.CommandButton CmdCancelar
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   6750
         TabIndex        =   31
         Top             =   1125
         Width           =   1860
      End
      Begin VB.CommandButton CmdHistorial
         Caption         =   "Ver Historial"
         Height          =   315
         Left            =   4680
         TabIndex        =   32
         Top             =   1125
         Width           =   1860
      End
      Begin VB.CommandButton CmdReg 
         Caption         =   "Registrar"
         Height          =   315
         Left            =   8775
         TabIndex        =   11
         Top             =   1125
         Width           =   1860
      End
      Begin VB.ComboBox CmbSta 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   270
         Width           =   3000
      End
      Begin VB.TextBox txtRemark 
         Height          =   315
         Left            =   900
         MaxLength       =   500
         TabIndex        =   8
         Top             =   630
         Width           =   9735
      End
      Begin VB.TextBox TxtRegist 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   900
         MaxLength       =   15
         TabIndex        =   9
         Top             =   1125
         Width           =   2985
      End
      Begin MSComCtl2.DTPicker DtpHor 
         Height          =   315
         Left            =   7695
         TabIndex        =   6
         Top             =   270
         Width           =   1770
         _ExtentX        =   3122
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60030978
         CurrentDate     =   40522
      End
      Begin MSComCtl2.DTPicker DtpFec 
         Height          =   315
         Left            =   4680
         TabIndex        =   5
         Top             =   270
         Width           =   2265
         _ExtentX        =   3995
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60030977
         CurrentDate     =   40522
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha:"
         Height          =   240
         Left            =   4095
         TabIndex        =   25
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label10 
         Caption         =   "Hora:"
         Height          =   240
         Left            =   7020
         TabIndex        =   24
         Top             =   315
         Width           =   510
      End
      Begin VB.Label Label4 
         Caption         =   "Status:"
         Height          =   240
         Left            =   180
         TabIndex        =   7
         Top             =   315
         Width           =   600
      End
      Begin VB.Line Line1 
         X1              =   180
         X2              =   10620
         Y1              =   1035
         Y2              =   1035
      End
      Begin VB.Label Label3 
         Caption         =   "Codigo:"
         Height          =   240
         Left            =   180
         TabIndex        =   3
         Top             =   1125
         Width           =   600
      End
      Begin VB.Label Label2 
         Caption         =   "Remark:"
         Height          =   240
         Left            =   180
         TabIndex        =   2
         Top             =   675
         Width           =   600
      End
   End
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   4200
      Left            =   2880
      TabIndex        =   0
      Top             =   1755
      Width           =   7395
      _ExtentX        =   13044
      _ExtentY        =   7408
      _Version        =   393216
      FixedCols       =   0
      SelectionMode   =   1
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   10
      Top             =   6015
      Width           =   13785
      _ExtentX        =   24315
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
Attribute VB_Name = "FrmConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Confirmaciones() As DatosConfirmacion
Private Cantidad As Long
Private Registrando As Boolean
Private Reiniciando As Boolean
Private CodigoHistorial As String
Private EstadoHistorial As String
Private FechaHistorial As Date
Private HistorialPendiente As Boolean
Private Resultado As String

Private Sub Form_Load()
    Dim Panel As MSComctlLib.Panel
    On Error GoTo ErrInt
    Me.Icon = MDIPri.ICO.ListImages(25).Picture
    ViewStatus.ImageList = MDIPri.ICO
    For Each Panel In StatusBar1.Panels: Panel.MinWidth = 15: Next
    StatusBar1.Panels(1).Width = 3000: StatusBar1.Panels(2).AutoSize = sbrSpring: StatusBar1.Panels(3).Width = 2700
    Grid.Top = 1800: Grid.Height = 4155
    Grid.Rows = 1: Grid.Cols = 5
    Grid.TextMatrix(0, 0) = "Codigo": Grid.ColWidth(0) = 900
    Grid.TextMatrix(0, 1) = "Empresa": Grid.ColWidth(1) = 1500
    Grid.TextMatrix(0, 2) = "Status": Grid.ColWidth(2) = 1200
    Grid.TextMatrix(0, 3) = "Fecha": Grid.ColWidth(3) = 1700
    Grid.TextMatrix(0, 4) = "Remark": Grid.ColWidth(4) = 1600
    CargarLista CmbSta, "SELECT ITEM,Status FROM dbo.IntStatus WHERE Active=1 AND ITEM>1 ORDER BY ITEM"
    TxtRegist.MaxLength = 18: txtRemark.MaxLength = 500
    TxtRegist.ToolTipText = "Enter agrega la guia al lote. Suprimir retira una fila pendiente."
    DtpHor.CheckBox = True
    DtpHor.ToolTipText = "Marque la casilla e indique la hora. Medianoche (00:00:00) es valida."
    LimpiarPendientes
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub ActualizarEstado()
    If Reiniciando Then Exit Sub
    CmdReg.Enabled = (Cantidad > 0 And Not Registrando)
    CmdCancelar.Enabled = Not Registrando
    CmdHistorial.Enabled = (Len(CodigoHistorial) > 0 And Not Registrando)
    CmbSta.Enabled = Not Registrando: DtpFec.Enabled = Not Registrando: DtpHor.Enabled = Not Registrando
    txtRemark.Enabled = Not Registrando: TxtRegist.Enabled = Not Registrando
    If Registrando Then
        StatusBar1.Panels(1).Text = "Registrando confirmaciones"
    ElseIf HistorialPendiente Then
        StatusBar1.Panels(1).Text = "Historial sin actualizar"
    ElseIf Len(Resultado) > 0 Then
        StatusBar1.Panels(1).Text = Resultado
    Else
        StatusBar1.Panels(1).Text = "Enter: agregar | Supr: retirar"
    End If
    If Len(CodigoHistorial) > 0 Then StatusBar1.Panels(2).Text = "Historial de guia: " & CodigoHistorial Else StatusBar1.Panels(2).Text = "Seleccione estado, fecha y hora"
    StatusBar1.Panels(3).Text = "Confirmaciones pendientes: " & CStr(Cantidad)
End Sub

Private Sub LimpiarPendientes()
    Reiniciando = True
    Cantidad = 0: ReDim Confirmaciones(1 To 1)
    Grid.Rows = 1: LblContReg.Caption = "0"
    CmbSta.ListIndex = -1: DtpFec.Value = Date: DtpHor.Value = Null
    txtRemark.Text = "": TxtRegist.Text = ""
    ViewStatus.Nodes.Clear: CodigoHistorial = "": EstadoHistorial = "": FechaHistorial = 0: HistorialPendiente = False
    Reiniciando = False
    ActualizarEstado
End Sub

Private Sub ConsultarHistorial(ByVal Codigo As String, Optional ByVal Estado As String = "", Optional ByVal Fecha As Date = 0, Optional ByVal Mensaje As String = "")
    On Error GoTo ErrInt
    CodigoHistorial = Codigo: EstadoHistorial = Estado: FechaHistorial = Fecha
    MostrarHistorial Codigo, ViewStatus, Estado, Fecha
    HistorialPendiente = False: ActualizarEstado
    Exit Sub
ErrInt:
    ViewStatus.Nodes.Clear: HistorialPendiente = True
    ActualizarEstado
    If Len(Mensaje) = 0 Then Mensaje = "No se pudo actualizar la vista del historial."
    MsgBox Mensaje & vbCrLf & Err.Description & vbCrLf & "Use Ver Historial para volver a consultarlo.", vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdHistorial_Click()
    If Not CmdHistorial.Enabled Then Exit Sub
    ConsultarHistorial CodigoHistorial, EstadoHistorial, FechaHistorial
End Sub

Public Sub AgregarConfirmacion(ByVal Codigo As String)
    Dim Datos As ADODB.Recordset, Posicion As Long, Fecha As Date, Empresa As String
    If Registrando Then Exit Sub
    If Not CodigoValido(Codigo) Then Err.Raise 5, , "Ingrese un codigo de documento valido."
    If CmbSta.ListIndex < 0 Then Err.Raise 5, , "Seleccione el estado de confirmacion."
    If IsNull(DtpHor.Value) Then Err.Raise 5, , "Marque la casilla de hora e indique la hora de confirmacion."
    Fecha = DateValue(DtpFec.Value) + TimeSerial(Hour(DtpHor.Value), Minute(DtpHor.Value), Second(DtpHor.Value))
    If Fecha < DateSerial(1753, 1, 1) Then Err.Raise 5, , "La fecha no es valida para la base de datos."
    Codigo = CStr(CDec(Codigo))
    For Posicion = 1 To Cantidad
        If Confirmaciones(Posicion).Codigo = Codigo Then Err.Raise 5, , "El documento ya existe en el listado."
    Next
    Set Datos = ConsultarDatos("SELECT NomEmp FROM dbo.ViewDatosDocumentos WHERE NroCod=? AND Active=1", Codigo, True)
    If Datos.EOF Then Datos.Close: Err.Raise 5, , "No se encontro el documento o fue anulado."
    Empresa = CStr(Datos.Fields(0).Value): Datos.Close
    ReDim Preserve Confirmaciones(1 To Cantidad + 1)
    Cantidad = Cantidad + 1
    With Confirmaciones(Cantidad)
        .Codigo = Codigo: .Estado = CmbSta.ItemData(CmbSta.ListIndex): .Fecha = Fecha: .Informe = txtRemark.Text
    End With
    Grid.Rows = Cantidad + 1
    Grid.TextMatrix(Cantidad, 0) = Codigo: Grid.TextMatrix(Cantidad, 1) = Empresa
    Grid.TextMatrix(Cantidad, 2) = CmbSta.Text: Grid.TextMatrix(Cantidad, 3) = Format$(Fecha, "dd/mm/yyyy hh:nn:ss")
    Grid.TextMatrix(Cantidad, 4) = txtRemark.Text
    LblContReg.Caption = CStr(Cantidad): TxtRegist.Text = "": Resultado = "Agregada al lote sin guardar"
    ActualizarEstado
    ConsultarHistorial Codigo, CmbSta.Text, Fecha, "La confirmacion ya esta en la lista pendiente, aun sin guardar. No repita el ingreso. Solo fallo la consulta del historial."
End Sub

Public Sub GuardarPendientes()
    Dim Codigo As String, Numero As Long, Descripcion As String, Total As Long
    On Error GoTo ErrInt
    If Registrando Then Exit Sub
    If Cantidad = 0 Then Err.Raise 5, , "No existen confirmaciones para registrar."
    If Len(Trim$(TxtRegist.Text)) > 0 Then Err.Raise 5, , "Hay un codigo sin agregar con Enter. Agreguelo o borrelo antes de registrar el lote."
    Registrando = True: ActualizarEstado
    MdiEstados.GuardarConfirmaciones Confirmaciones, Cantidad
    Codigo = Confirmaciones(Cantidad).Codigo: Total = Cantidad
    LimpiarPendientes
    Resultado = CStr(Total) & " confirmaciones registradas"
    Registrando = False: ActualizarEstado
    ConsultarHistorial Codigo, , , "El lote de " & CStr(Total) & " confirmaciones ya esta guardado. No vuelva a registrarlo. Solo fallo la consulta del historial."
    Exit Sub
ErrInt:
    Numero = Err.Number: Descripcion = Err.Description
    Registrando = False: ActualizarEstado
    Err.Raise Numero, "GuardarPendientes", Descripcion
End Sub

Public Sub QuitarConfirmacion(ByVal Fila As Long)
    Dim Posicion As Long
    If Registrando Or Fila < 1 Or Fila > Cantidad Then Exit Sub
    For Posicion = Fila To Cantidad - 1: Confirmaciones(Posicion) = Confirmaciones(Posicion + 1): Next
    If Cantidad = 1 Then Grid.Rows = 1 Else Grid.RemoveItem Fila
    Cantidad = Cantidad - 1: LblContReg.Caption = CStr(Cantidad)
    ViewStatus.Nodes.Clear: CodigoHistorial = "": EstadoHistorial = "": HistorialPendiente = False
    Resultado = "Confirmacion retirada del lote": ActualizarEstado
End Sub

Private Sub CmdReg_Click()
    On Error GoTo ErrInt
    If Not CmdReg.Enabled Then Exit Sub
    If Len(Trim$(TxtRegist.Text)) > 0 Then Err.Raise 5, , "Hay un codigo sin agregar con Enter. Agreguelo o borrelo antes de registrar el lote."
    If MsgBox("Se registraran " & CStr(Cantidad) & " confirmaciones del listado. Desea continuar?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") <> vbYes Then Exit Sub
    GuardarPendientes
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdCancelar_Click()
    If PuedeCerrar() Then LimpiarPendientes
End Sub

Private Sub TxtRegist_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrInt
    Select Case KeyAscii
        Case 13: KeyAscii = 0: AgregarConfirmacion Trim$(TxtRegist.Text)
        Case 48 To 57, 8, 22
        Case Else: KeyAscii = 0
    End Select
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub Grid_Click()
    If Grid.Row < 1 Or Grid.Row > Cantidad Or Registrando Then Exit Sub
    ConsultarHistorial Confirmaciones(Grid.Row).Codigo, Grid.TextMatrix(Grid.Row, 2), Confirmaciones(Grid.Row).Fecha
End Sub

Private Sub Grid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyDelete Or Grid.Row < 1 Or Cantidad = 0 Or Registrando Then Exit Sub
    KeyCode = 0
    If MsgBox("Desea quitar la confirmacion de la guia " & Grid.TextMatrix(Grid.Row, 0) & " de la lista pendiente? No se borrara su historial.", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") = vbYes Then QuitarConfirmacion Grid.Row
End Sub

Private Sub txtRemark_Change()
    Dim Posicion As Long
    Posicion = txtRemark.SelStart
    If txtRemark.Text <> UCase$(txtRemark.Text) Then txtRemark.Text = UCase$(txtRemark.Text): txtRemark.SelStart = Posicion
End Sub

Private Sub txtRemark_GotFocus()
    txtRemark.SelStart = 0: txtRemark.SelLength = Len(txtRemark.Text)
End Sub

Public Function PuedeCerrar() As Boolean
    If Registrando Then Exit Function
    PuedeCerrar = True
    If Cantidad = 0 And Len(TxtRegist.Text & txtRemark.Text) = 0 And CmbSta.ListIndex < 0 And IsNull(DtpHor.Value) And DateValue(DtpFec.Value) = Date Then Exit Function
    PuedeCerrar = (MsgBox("Hay confirmaciones o campos en borrador sin registrar. Desea descartarlos?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") = vbYes)
    If PuedeCerrar Then Resultado = "Datos pendientes descartados": LimpiarPendientes
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then KeyCode = 0: CmdCancelar_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Not CerrandoAplicacion Then Cancel = Not PuedeCerrar()
End Sub

Private Sub Form_Activate()
    VerificarAccesoFormulario Me, 7
End Sub
