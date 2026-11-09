VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmPermisos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Configuracion de Accesos"
   ClientHeight    =   7755
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7815
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7755
   ScaleWidth      =   7815
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox CmbUsu 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5880
      TabIndex        =   6
      Text            =   "ADMIN"
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton CmdPri
      Height = 375
      Left = 120
      Top = 120
      Width = 2295
      TabIndex = 0
      Caption = "Marcar Todos"
   End
   Begin VB.CommandButton CmdMod
      Height = 375
      Left = 2400
      Top = 120
      Width = 2295
      TabIndex = 1
      Caption = "Desmarcar Todos"
   End
   Begin MSComctlLib.TreeView ViewMenus 
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   11033
      _Version        =   393217
      Indentation     =   90
      LineStyle       =   1
      Style           =   7
      Checkboxes      =   -1  'True
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   330
      Left            =   0
      TabIndex        =   3
      Top             =   7425
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "23/06/2010"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   17639
            MinWidth        =   17639
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton CmdAceptar
      Height = 375
      Left = 6000
      Top = 6960
      Width = 1695
      TabIndex = 4
      Caption = "Guardar Permisos"
   End
   Begin VB.CommandButton CmdCancelar
      Caption = "Cancelar"
      Left = 3600
      Top = 6960
      Width = 2200
      Height = 375
      TabIndex = 7
   End
   Begin VB.Label Label1 
      Caption = "Permisos por modulo"
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   6960
      Width           =   3000
   End
End
Attribute VB_Name = "FrmPermisos"
Attribute VB_PredeclaredId = True
Option Explicit
Private Codigo As Long
Private Cambio As Boolean
Private Administrador As Boolean

Public Sub Preparar(ByVal Usuario As Long, ByVal Nombre As String, ByVal EsAdministrador As Boolean)
    Dim Modulo As Integer, Datos As ADODB.Recordset
    ExigirAdministrador
    Codigo = Usuario: CmbUsu.Text = Nombre: CmbUsu.Locked = True: Administrador = EsAdministrador
    ViewMenus.Nodes.Clear
    For Modulo = 1 To 11
        ViewMenus.Nodes.Add , , "M" & CStr(Modulo), NombreModulo(Modulo)
        ViewMenus.Nodes("M" & CStr(Modulo)).Checked = EsAdministrador
    Next
    Set Datos = ConsultarDatos("SELECT CodMod,Permitido FROM dbo.IntAccesos WHERE CodUsu=?", CStr(Usuario), True)
    Do Until Datos.EOF
        If Not Administrador Then ViewMenus.Nodes("M" & CStr(Datos!CodMod)).Checked = CBool(Datos!Permitido)
        Datos.MoveNext
    Loop
    Datos.Close: Cambio = False
    ViewMenus.Enabled = Not Administrador: CmdPri.Enabled = Not Administrador: CmdMod.Enabled = Not Administrador: CmdAceptar.Enabled = Not Administrador
    Dim Panel As MSComctlLib.Panel
    For Each Panel In StatusBar1.Panels: Panel.MinWidth = 15: Panel.Style = sbrText: Panel.Width = 15: Next
    StatusBar1.Panels(1).AutoSize = sbrSpring
    If Administrador Then StatusBar1.Panels(1).Text = "Administrador: acceso a todos los modulos" Else StatusBar1.Panels(1).Text = "Marque los modulos y presione Guardar"
End Sub

Private Sub CmdPri_Click()
    Dim Nodo As MSComctlLib.Node
    For Each Nodo In ViewMenus.Nodes: Nodo.Checked = True: Next
    Cambio = True
End Sub

Private Sub CmdMod_Click()
    Dim Nodo As MSComctlLib.Node
    For Each Nodo In ViewMenus.Nodes: Nodo.Checked = False: Next
    Cambio = True
End Sub

Private Sub ViewMenus_NodeCheck(ByVal Node As MSComctlLib.Node)
    Cambio = True
End Sub

Private Sub CmdAceptar_Click()
    Dim Lista As String, Modulo As Integer
    On Error GoTo ErrorGuardar
    If Administrador Then Exit Sub
    Lista = "<Permisos>"
    For Modulo = 1 To 11
        Lista = Lista & "<Modulo Codigo=""" & CStr(Modulo) & """ Permitir=""" & IIf(ViewMenus.Nodes("M" & CStr(Modulo)).Checked, "1", "0") & """/>"
    Next
    Lista = Lista & "</Permisos>"
    GuardarUsuario Codigo, 5, , , , , , , , , Lista
    Cambio = False: Unload Me
    Exit Sub
ErrorGuardar:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Cambio Then Cancel = (MsgBox("Hay permisos sin guardar. Desea descartarlos?", vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") <> vbYes)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then KeyCode = 0: Unload Me
End Sub

Private Sub Form_Activate()
    VerificarAccesoFormulario Me, 0
End Sub
