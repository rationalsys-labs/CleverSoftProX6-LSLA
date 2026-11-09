VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmBusqueda 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Busqueda de Cliente"
   ClientHeight    =   5940
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7500
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5940
   ScaleWidth      =   7500
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid GridBus 
      Height          =   5550
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5820
      _ExtentX        =   10266
      _ExtentY        =   9790
      _Version        =   393216
      FixedCols       =   0
      Appearance      =   0
   End
End
Attribute VB_Name = "FrmBusqueda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public VariableBusCli As String
Private Aceptado As Boolean

Private Sub Form_Load()
    Dim Columna As Integer
    On Error GoTo ErrInt
    Aceptado = False
    For Columna = 0 To 4: Var.ResulBus(Columna) = vbNullString: Next
    Me.Height = GridBus.Height
    Me.Width = GridBus.Width - 50
    GridBus.Cols = 5: GridBus.Rows = 1
    GridBus.TextMatrix(0, 0) = "Codigo": GridBus.ColWidth(0) = 800
    GridBus.TextMatrix(0, 1) = "Cliente": GridBus.ColWidth(1) = 2800
    GridBus.TextMatrix(0, 2) = "Documento": GridBus.ColWidth(2) = 1200
    GridBus.ColWidth(3) = 0: GridBus.ColWidth(4) = 0
    CargarBusqueda
    Exit Sub
ErrInt:
    MsgBox "No se pudo buscar clientes: " & Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Public Sub CargarBusqueda()
    Dim Datos As ADODB.Recordset
    Dim Fila As Long
    Dim Columna As Integer
    Set Datos = BuscarClientes(VariableBusCli)
    GridBus.Rows = 1
    Do Until Datos.EOF
        GridBus.Rows = GridBus.Rows + 1
        Fila = GridBus.Rows - 1
        For Columna = 0 To 4: GridBus.TextMatrix(Fila, Columna) = TextoDato(Datos.Fields(Columna).Value): Next
        Datos.MoveNext
    Loop
    Datos.Close
End Sub

Public Sub AceptarCliente()
    Dim Columna As Integer
    If GridBus.Row < 1 Or GridBus.Rows < 2 Then Exit Sub
    Aceptado = False
    For Columna = 0 To 4: Var.ResulBus(Columna) = GridBus.TextMatrix(GridBus.Row, Columna): Next
    Aceptado = True
    Unload Me
End Sub

Private Sub GridBus_DblClick()
    AceptarCliente
End Sub

Private Sub GridBus_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: AceptarCliente
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then KeyCode = 0: Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim Columna As Integer
    If Not Aceptado Then
        Aceptado = False
    For Columna = 0 To 4: Var.ResulBus(Columna) = vbNullString: Next
    End If
End Sub
