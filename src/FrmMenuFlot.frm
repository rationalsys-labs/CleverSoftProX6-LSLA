VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FrmMenuFlot 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2505
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9525
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2505
   ScaleWidth      =   9525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid Grid 
      Height          =   2355
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   4154
      _Version        =   393216
      FixedCols       =   0
      Appearance      =   0
   End
   Begin VB.CommandButton CmdFiltroGrupo
      Caption         =   "Buscar (F3)"
      Height          =   300
      Left            =   0
      Top             =   2400
      Width           =   1500
      TabIndex        =   2
      Visible         =   0
   End
   Begin VB.Label LblAyudaGrupo
      Height          =   300
      Left            =   1575
      Top             =   2430
      Width           =   7650
      TabIndex        =   3
      Visible         =   0
   End
   Begin VB.Label Lblcol 
      Height          =   150
      Left            =   1530
      TabIndex        =   1
      Top             =   2340
      Visible         =   0   'False
      Width           =   870
   End
End
Attribute VB_Name = "FrmMenuFlot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ListaLimitada As Boolean

Private Sub Form_Load()
    Dim Columna As Integer
    On Error GoTo ErrInt
    Me.Icon = MDIPri.ICO.ListImages(25).Picture
    Me.Height = Grid.Height: Me.Width = Grid.Width
    Var.RegistroSeleccionado = False
    For Columna = 0 To 5: Var.ResultadoRegistro(Columna) = vbNullString: Next
    Grid.Rows = 1
    Select Case Var.Varflotante
        Case 1
            Grid.Cols = 6
            Grid.TextMatrix(0, 0) = "Codigo": Grid.ColWidth(0) = 1100
            Grid.TextMatrix(0, 1) = "Boleta": Grid.ColWidth(1) = 1100
            Grid.TextMatrix(0, 2) = "Cliente": Grid.ColWidth(2) = 4500
            Grid.TextMatrix(0, 3) = "CodCliente": Grid.ColWidth(3) = 0
            Grid.TextMatrix(0, 4) = "Producto": Grid.ColWidth(4) = 1500
            Grid.TextMatrix(0, 5) = "Cantidad": Grid.ColWidth(5) = 800
            Lblcol.Caption = "Cliente"
        Case 2
            Grid.Cols = 5
            Grid.TextMatrix(0, 0) = "Codigo": Grid.ColWidth(0) = 1100
            Grid.TextMatrix(0, 1) = "Empresa": Grid.ColWidth(1) = 3700
            Grid.TextMatrix(0, 2) = "Direccion2": Grid.ColWidth(2) = 2000
            Grid.TextMatrix(0, 3) = "Contacto": Grid.ColWidth(3) = 1500
            Grid.TextMatrix(0, 4) = "Servicio": Grid.ColWidth(4) = 700
            Lblcol.Caption = "Empresa"
        Case 3
            Grid.Cols = 2
            Grid.TextMatrix(0, 0) = "Item": Grid.ColWidth(0) = 700
            Grid.TextMatrix(0, 1) = "NomGrp": Grid.ColWidth(1) = 8200
            Lblcol.Caption = "NomGrp"
            Me.Height = Grid.Height + 390
            CmdFiltroGrupo.Visible = True: LblAyudaGrupo.Visible = True
    End Select
    FiltrarRegistros vbNullString
    Exit Sub
ErrInt:
    MsgBox "No se pudieron cargar los registros: " & Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Public Sub FiltrarRegistros(ByVal Texto As String)
    Dim Datos As ADODB.Recordset
    Dim Fila As Long
    Dim Columna As Integer
    Set Datos = BuscarRegistros(Var.Varflotante, Lblcol.Caption, Texto)
    Grid.Rows = 1: ListaLimitada = False
    Do Until Datos.EOF
        If Var.Varflotante = 3 And Fila = 50 Then ListaLimitada = True: Exit Do
        Grid.Rows = Grid.Rows + 1: Fila = Grid.Rows - 1
        For Columna = 0 To Grid.Cols - 1: Grid.TextMatrix(Fila, Columna) = TextoDato(Datos.Fields(Columna).Value): Next
        Datos.MoveNext
    Loop
    Datos.Close
    ActualizarAyudaGrupo
End Sub

Public Sub AceptarRegistro()
    Dim Columna As Integer
    If Grid.Row < 1 Or Grid.Rows < 2 Then Exit Sub
    For Columna = 0 To Grid.Cols - 1: Var.ResultadoRegistro(Columna) = Grid.TextMatrix(Grid.Row, Columna): Next
    Var.RegistroSeleccionado = True
    Unload Me
End Sub

Private Sub Grid_Click()
    If Var.Varflotante = 1 And Grid.Row > 0 Then
        AceptarRegistro
    Else
        Lblcol.Caption = Grid.TextMatrix(0, Grid.Col)
        ActualizarAyudaGrupo
    End If
End Sub

Private Sub Grid_DblClick()
    AceptarRegistro
End Sub

Private Sub Grid_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then KeyAscii = 0: AceptarRegistro
End Sub

Private Sub Grid_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = vbRightButton And Var.Varflotante <> 1 Then Me.PopupMenu MDIPri.MenuBus
End Sub

Private Sub ActualizarAyudaGrupo()
    Dim Campo As String
    If Var.Varflotante <> 3 Then Exit Sub
    If Lblcol.Caption = "Item" Then Campo = "codigo" Else Campo = "nombre"
    If ListaLimitada Then
        LblAyudaGrupo.Caption = "Primeros 50 grupos. Use Buscar (F3) por " & Campo & " para precisar."
    Else
        LblAyudaGrupo.Caption = CStr(Grid.Rows - 1) & " grupos. Buscar (F3) por " & Campo & "; clic en columna para cambiar."
    End If
End Sub

Private Sub CmdFiltroGrupo_Click()
    On Error GoTo ErrInt
    If Var.Varflotante <> 3 Then Exit Sub
    Load FrmSerch
    If Lblcol.Caption = "Item" Then FrmSerch.Caption = "Buscar Grupo por Codigo" Else FrmSerch.Caption = "Buscar Grupo por Nombre"
    FrmSerch.Show vbModal, Me
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 And Var.Varflotante = 3 Then KeyCode = 0: CmdFiltroGrupo_Click
    If KeyCode = vbKeyEscape Then KeyCode = 0: Unload Me
End Sub
