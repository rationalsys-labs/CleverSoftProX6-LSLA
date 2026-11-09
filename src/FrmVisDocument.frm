VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVisDocument 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Visualizar Informacion de registros"
   ClientHeight    =   5250
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   17970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   17970
   Begin VB.CommandButton CmdBuscar 
      Height          =   315
      Left            =   5760
      Picture         =   "FrmVisDocument.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   40
      Top             =   120
      Width           =   420
   End
   Begin VB.TextBox TxtRegist 
      Height          =   315
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Width           =   2775
   End
   Begin MSComctlLib.TreeView ViewHis 
      Height          =   4380
      Left            =   9630
      TabIndex        =   34
      Top             =   540
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   7726
      _Version        =   393217
      Indentation     =   18
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos del Documento"
      Height          =   2220
      Left            =   135
      TabIndex        =   20
      Top             =   2700
      Width           =   9375
      Begin VB.TextBox TxtDescrip 
         Height          =   285
         Left            =   7290
         TabIndex        =   42
         Top             =   1350
         Width           =   1860
      End
      Begin VB.TextBox TxtDocRef 
         Height          =   285
         Left            =   4320
         TabIndex        =   38
         Top             =   225
         Width           =   1860
      End
      Begin VB.TextBox TxtDocFec 
         Height          =   285
         Left            =   7290
         TabIndex        =   37
         Top             =   630
         Width           =   1860
      End
      Begin VB.TextBox TxtDocCon 
         Height          =   285
         Left            =   4275
         TabIndex        =   26
         Top             =   1710
         Width           =   4875
      End
      Begin VB.TextBox TxtDocDis 
         Height          =   285
         Left            =   1170
         TabIndex        =   25
         Top             =   1710
         Width           =   1860
      End
      Begin VB.TextBox TxtDocDir 
         Height          =   285
         Left            =   1170
         TabIndex        =   24
         Top             =   1350
         Width           =   5010
      End
      Begin VB.TextBox TxtDocPro 
         Height          =   285
         Left            =   1170
         TabIndex        =   23
         Top             =   630
         Width           =   5010
      End
      Begin VB.TextBox TxtDocDes 
         Height          =   285
         Left            =   1170
         TabIndex        =   22
         Top             =   990
         Width           =   7980
      End
      Begin VB.TextBox TxtDocCod 
         Height          =   285
         Left            =   1170
         TabIndex        =   21
         Top             =   270
         Width           =   1860
      End
      Begin VB.Label Label19 
         Caption         =   "Descripcion:"
         Height          =   195
         Left            =   6345
         TabIndex        =   43
         Top             =   1395
         Width           =   915
      End
      Begin VB.Label Label17 
         Caption         =   "Cod Cliente:"
         Height          =   195
         Left            =   3375
         TabIndex        =   39
         Top             =   270
         Width           =   960
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   6345
         TabIndex        =   36
         Top             =   675
         Width           =   645
      End
      Begin VB.Label Label12 
         Caption         =   "Contacto:"
         Height          =   195
         Left            =   3240
         TabIndex        =   32
         Top             =   1755
         Width           =   1005
      End
      Begin VB.Label Label13 
         Caption         =   "Distrito:"
         Height          =   195
         Left            =   225
         TabIndex        =   31
         Top             =   1755
         Width           =   645
      End
      Begin VB.Label Label14 
         Caption         =   "Direccion:"
         Height          =   195
         Left            =   225
         TabIndex        =   30
         Top             =   1395
         Width           =   780
      End
      Begin VB.Label Label15 
         Caption         =   "Producto:"
         Height          =   195
         Left            =   225
         TabIndex        =   29
         Top             =   675
         Width           =   915
      End
      Begin VB.Label Label16 
         Caption         =   "Destinat:"
         Height          =   195
         Left            =   225
         TabIndex        =   28
         Top             =   1035
         Width           =   645
      End
      Begin VB.Label Label18 
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   225
         TabIndex        =   27
         Top             =   315
         Width           =   645
      End
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   240
      Left            =   0
      TabIndex        =   33
      Top             =   5010
      Width           =   17970
      _ExtentX        =   31697
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
            Object.Width           =   26459
            MinWidth        =   26459
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Servicio"
      Height          =   2220
      Left            =   135
      TabIndex        =   1
      Top             =   450
      Width           =   9375
      Begin VB.CommandButton CmdReaServ 
         Caption         =   "Reactivar Servicio"
         Enabled         =   0   'False
         Height          =   330
         Left            =   7290
         TabIndex        =   41
         Top             =   225
         Width           =   1860
      End
      Begin VB.TextBox TxtServCan 
         Height          =   285
         Left            =   1170
         TabIndex        =   19
         Top             =   1710
         Width           =   1860
      End
      Begin VB.TextBox TxtServPro 
         Height          =   285
         Left            =   7290
         TabIndex        =   17
         Top             =   1350
         Width           =   1860
      End
      Begin VB.TextBox TxtServRep 
         Height          =   285
         Left            =   4275
         TabIndex        =   15
         Top             =   1350
         Width           =   1860
      End
      Begin VB.TextBox TxtServDis 
         Height          =   285
         Left            =   1170
         TabIndex        =   13
         Top             =   1350
         Width           =   1860
      End
      Begin VB.TextBox TxtServDir 
         Height          =   285
         Left            =   1170
         TabIndex        =   11
         Top             =   990
         Width           =   7980
      End
      Begin VB.TextBox TxtServDoc 
         Height          =   285
         Left            =   7290
         TabIndex        =   9
         Top             =   630
         Width           =   1860
      End
      Begin VB.TextBox TxtServCli 
         Height          =   285
         Left            =   1170
         TabIndex        =   7
         Top             =   630
         Width           =   4965
      End
      Begin VB.TextBox TxtServBol 
         Height          =   285
         Left            =   4275
         TabIndex        =   5
         Top             =   270
         Width           =   1860
      End
      Begin VB.TextBox TxtServCod 
         Height          =   285
         Left            =   1170
         TabIndex        =   3
         Top             =   270
         Width           =   1860
      End
      Begin VB.Label Label9 
         Caption         =   "Cantidad:"
         Height          =   195
         Left            =   225
         TabIndex        =   18
         Top             =   1755
         Width           =   645
      End
      Begin VB.Label Label8 
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   6345
         TabIndex        =   16
         Top             =   1395
         Width           =   645
      End
      Begin VB.Label Label7 
         Caption         =   "Tipo de Serv:"
         Height          =   195
         Left            =   3240
         TabIndex        =   14
         Top             =   1395
         Width           =   1005
      End
      Begin VB.Label Label6 
         Caption         =   "Distrito:"
         Height          =   195
         Left            =   225
         TabIndex        =   12
         Top             =   1395
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Direccion:"
         Height          =   195
         Left            =   225
         TabIndex        =   10
         Top             =   1035
         Width           =   780
      End
      Begin VB.Label Label4 
         Caption         =   "Documento:"
         Height          =   195
         Left            =   6345
         TabIndex        =   8
         Top             =   675
         Width           =   915
      End
      Begin VB.Label Label3 
         Caption         =   "Cliente:"
         Height          =   195
         Left            =   225
         TabIndex        =   6
         Top             =   675
         Width           =   645
      End
      Begin VB.Label Label2 
         Caption         =   "Boleta:"
         Height          =   195
         Left            =   3285
         TabIndex        =   4
         Top             =   315
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo:"
         Height          =   195
         Left            =   225
         TabIndex        =   2
         Top             =   315
         Width           =   645
      End
   End
   Begin VB.Label Label10 
      Caption         =   "Buscar Por Codigo de Documento:"
      Height          =   195
      Left            =   180
      TabIndex        =   35
      Top             =   135
      Width           =   2640
   End
End
Attribute VB_Name = "FrmVisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Dim Control As Control
    Me.Icon = MDIPri.ICO.ListImages(25).Picture
    ViewHis.ImageList = MDIPri.ICO
    For Each Control In Me.Controls
        If TypeOf Control Is TextBox Then
            If Control.Name <> "TxtRegist" Then Control.Locked = True
        End If
    Next
End Sub

Public Sub CargarRegistro(ByVal Codigo As String)
    Dim Datos As ADODB.Recordset
    Dim Numero As Long
    Dim PuedeReactivar As Boolean
    Dim NumeroError As Long, Descripcion As String
    On Error GoTo ErrInt
    LimpiarConsulta
    If Not CodigoValido(Codigo) Then Err.Raise vbObjectError + 350, , "Especifique un codigo de registro valido."
    Set Datos = ConsultarDatos("SELECT Servicio,Boleta,Cliente,RucDni,Direccion,Distrito,Reparto,Producto,Cantidad,Codigo,ProdCliente,Fecha,Empresa,Direccion2,Distrito2,Contacto,CodCliente,Estado,Descripcion FROM dbo.ViewVisDocument WHERE Codigo=?", Codigo, True)
    If Datos.EOF Then Datos.Close: Err.Raise vbObjectError + 351, , "No se ubica el registro."
    TxtServCod.Text = CStr(Datos.Fields(0).value): TxtServBol.Text = CStr(Datos.Fields(1).value)
    TxtServCli.Text = TextoDato(Datos.Fields(2).value): TxtServDoc.Text = TextoDato(Datos.Fields(3).value)
    TxtServDir.Text = TextoDato(Datos.Fields(4).value): TxtServDis.Text = TextoDato(Datos.Fields(5).value)
    TxtServRep.Text = TextoDato(Datos.Fields(6).value): TxtServPro.Text = TextoDato(Datos.Fields(7).value)
    TxtServCan.Text = TextoDato(Datos.Fields(8).value): TxtDocCod.Text = CStr(Datos.Fields(9).value)
    TxtDocPro.Text = TextoDato(Datos.Fields(10).value): TxtDocFec.Text = Format$(Datos.Fields(11).value, "dd/mm/yyyy hh:nn:ss")
    TxtDocDes.Text = TextoDato(Datos.Fields(12).value): TxtDocDir.Text = TextoDato(Datos.Fields(13).value)
    TxtDocDis.Text = TextoDato(Datos.Fields(14).value): TxtDocCon.Text = TextoDato(Datos.Fields(15).value)
    TxtDocRef.Text = TextoDato(Datos.Fields(16).value): PuedeReactivar = Not CBool(Datos.Fields(17).value)
    TxtDescrip.Text = TextoDato(Datos.Fields(18).value)
    Datos.Close
    Set Datos = ConsultarDatos("SELECT Status,Emision,Usuario,Ranche,Informe FROM dbo.ViewHistory WHERE Codigo=? ORDER BY Emision,Item", Codigo, True)
    ViewHis.Nodes.Add , , "Titulo", "CODIGO-" & Codigo, MDIPri.ICO.ListImages(23).Key
    Do Until Datos.EOF
        Numero = Numero + 1
        ViewHis.Nodes.Add "Titulo", tvwChild, "Evento" & CStr(Numero), TextoDato(Datos.Fields(0).value) & " - " & Format$(Datos.Fields(1).value, "dd/mm/yyyy hh:nn:ss") & " " & TextoDato(Datos.Fields(2).value) & " DETALLE: " & TextoDato(Datos.Fields(4).value), MDIPri.ICO.ListImages(24).Key
        Datos.MoveNext
    Loop
    ViewHis.Nodes("Titulo").Expanded = True
    Datos.Close
    CmdReaServ.Enabled = PuedeReactivar
    Exit Sub
ErrInt:
    NumeroError = Err.Number: Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then
        If Datos.State <> adStateClosed Then Datos.Close
    End If
    LimpiarConsulta
    On Error GoTo 0
    Err.Raise NumeroError, "CargarRegistro", Descripcion
End Sub

Private Sub LimpiarConsulta()
    Dim Control As Control
    For Each Control In Me.Controls
        If TypeOf Control Is TextBox Then
            If Control.Name <> "TxtRegist" Then Control.Text = ""
        End If
    Next
    ViewHis.Nodes.Clear: CmdReaServ.Enabled = False
End Sub

Private Sub Label10_Click()

End Sub

Private Sub TxtRegist_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrInt
    If KeyAscii <> 13 Then Exit Sub
    KeyAscii = 0
    CargarRegistro Trim$(TxtRegist.Text)
    TxtRegist.Text = ""
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdBuscar_Click()
    On Error GoTo ErrInt
    Var.Varflotante = 2
    FrmMenuFlot.Top = Me.Top + 1550: FrmMenuFlot.Left = Me.Left + 260
    FrmMenuFlot.Show vbModal
    If Var.RegistroSeleccionado Then CargarRegistro Var.ResultadoRegistro(0)
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdReaServ_Click()
    On Error GoTo ErrInt
    If Not CodigoValido(TxtServCod.Text) Then Exit Sub
    If MsgBox("Desea reactivar el servicio " & TxtServCod.Text & "?", vbQuestion Or vbYesNo, "Reactivar Servicio") <> vbYes Then Exit Sub
    CambiarServicio TxtServCod.Text, 3
    CargarRegistro TxtDocCod.Text
    Exit Sub
ErrInt:
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub Form_Activate()
    VerificarAccesoFormulario Me, 2
End Sub
