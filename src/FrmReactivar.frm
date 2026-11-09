VERSION 5.00
Begin VB.Form FrmReactivar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Liberar Guia"
   ClientHeight    =   1170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2670
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1170
   ScaleWidth      =   2670
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdreactivar 
      Caption         =   "Liberar"
      Height          =   375
      Left            =   135
      TabIndex        =   1
      Top             =   630
      Width           =   1150
   End
   Begin VB.CommandButton CmdCancelar
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   1410
      TabIndex        =   2
      Top             =   630
      Width           =   1150
   End
   Begin VB.TextBox Txtliberar 
      Height          =   330
      Left            =   315
      TabIndex        =   0
      Top             =   180
      Width           =   2085
   End
End
Attribute VB_Name = "FrmReactivar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Icon = MDIPri.ICO.ListImages(25).Picture
    Txtliberar.MaxLength = 18
    Txtliberar.ToolTipText = "Codigo de la guia cuya asignacion de reparto se liberara."
End Sub

Public Sub PrepararLiberacion(ByVal Codigo As String)
    If Not CodigoValido(Trim$(Codigo)) Then Err.Raise 5, , "Ingrese el codigo de la guia."
    Txtliberar.Text = Trim$(Codigo): Txtliberar.Locked = True
End Sub

Private Sub cmdreactivar_Click()
    Dim Datos As ADODB.Recordset, Mensaje As String, Codigo As String, Cargo As String
    Dim Descripcion As String
    On Error GoTo ErrInt
    Codigo = Trim$(Txtliberar.Text)
    If Not CodigoValido(Codigo) Then Err.Raise 5, , "Ingrese el codigo de la guia."
    Set Datos = ConsultarDatos("SELECT d.NroCod,e.NomEmp,h.Ranche,h.FecEmi,ISNULL(m.NomMen,h.Informe) Mensajero FROM dbo.MDBDocument d JOIN dbo.MDBHisEmp e ON e.CodEmp=d.CodEmp CROSS APPLY (SELECT TOP 1 Ranche,FecEmi,CodMen,Informe FROM dbo.MDBHistory WHERE NroCod=d.NroCod AND Status=1 AND Active=1 ORDER BY FecEmi DESC,ITEM DESC) h LEFT JOIN dbo.Mdbmensajeros m ON m.CodMen=h.CodMen WHERE d.NroCod=? AND d.Active=1", Codigo, True)
    If Datos.EOF Then Err.Raise 5, , "La guia no existe, fue anulada o no tiene una asignacion activa de reparto."
    Cargo = TextoDato(Datos.Fields("Ranche").Value)
    If Cargo = "0" Then Cargo = "Consolidado (sin numero de cargo)"
    Mensaje = "Desea liberar la asignacion de reparto de esta guia?" & vbCrLf & vbCrLf & _
        "Guia: " & TextoDato(Datos.Fields("NroCod").Value) & vbCrLf & _
        "Destinatario: " & TextoDato(Datos.Fields("NomEmp").Value) & vbCrLf & _
        "Mensajero: " & TextoDato(Datos.Fields("Mensajero").Value) & vbCrLf & _
        "Cargo: " & Cargo & vbCrLf & "Fecha de reparto: " & Format$(Datos.Fields("FecEmi").Value, "dd/mm/yyyy hh:nn:ss") & vbCrLf & vbCrLf & _
        "Se desactivaran sus asignaciones vigentes de reparto. Se conservaran el documento y su historial."
    Datos.Close: Set Datos = Nothing
    If MsgBox(Mensaje, vbYesNo Or vbQuestion Or vbDefaultButton2, "Clever SoftPro") <> vbYes Then Exit Sub
    LiberarGuia Codigo
    MsgBox "Guia " & Codigo & " liberada. Para agregarla a la lista, vuelva al codigo de guia y presione Enter.", vbInformation, "Clever SoftPro"
    Unload Me
    Exit Sub
ErrInt:
    Descripcion = Err.Description
    On Error Resume Next
    If Not Datos Is Nothing Then Datos.Close
    On Error GoTo 0
    MsgBox Descripcion, vbExclamation, "Clever SoftPro"
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Txtliberar_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13: KeyAscii = 0: cmdreactivar_Click
        Case 48 To 57, 8, 22
        Case Else: KeyAscii = 0
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then KeyCode = 0: Unload Me
End Sub
