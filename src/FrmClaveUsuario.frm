VERSION 5.00
Begin VB.Form FrmClaveUsuario
 Caption = "Cambiar Clave de Usuario"
 BorderStyle = 1
 ClientWidth = 5400
 ClientHeight = 2100
 MaxButton = 0
 MinButton = 0
 KeyPreview = -1
 StartUpPosition = 2
 Begin VB.Label LblUsuario
  Left = 180
  Top = 180
  Width = 5000
  Height = 300
 End
 Begin VB.Label LblClave
  Caption = "Nueva clave:"
  Left = 180
  Top = 675
  Width = 1450
  Height = 300
 End
 Begin VB.Label LblConfirmar
  Caption = "Confirmar clave:"
  Left = 180
  Top = 1080
  Width = 1450
  Height = 300
 End
 Begin VB.TextBox TxtClave
  PasswordChar = "*"
  MaxLength = 150
  Left = 1750
  Top = 630
  Width = 3400
  Height = 315
 End
 Begin VB.TextBox TxtConfirmar
  PasswordChar = "*"
  MaxLength = 150
  Left = 1750
  Top = 1035
  Width = 3400
  Height = 315
 End
 Begin VB.CommandButton CmdGuardar
  Caption = "Guardar"
  Enabled = 0
  Left = 1750
  Top = 1575
  Width = 1600
  Height = 375
 End
 Begin VB.CommandButton CmdCancelar
  Caption = "Cancelar"
  Left = 3550
  Top = 1575
  Width = 1600
  Height = 375
 End
End
Attribute VB_Name = "FrmClaveUsuario"
Attribute VB_PredeclaredId = True
Option Explicit
Private Codigo As Long
Public Sub Preparar(ByVal Usuario As Long, ByVal Nombre As String)
 ExigirAdministrador
 Codigo = Usuario: LblUsuario.Caption = "Usuario: " & Nombre
End Sub
Private Sub CmdGuardar_Click()
 On Error GoTo ErrorGuardar
 GuardarUsuario Codigo, 4, , , , TxtClave.Text, TxtConfirmar.Text
 TxtClave.Text = "": TxtConfirmar.Text = ""
 Unload Me
 Exit Sub
ErrorGuardar:
 MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub
Private Sub CmdCancelar_Click()
 Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyEscape Then KeyCode = 0: Unload Me
End Sub
Private Sub Form_Activate()
    VerificarAccesoFormulario Me, 0
End Sub

Private Sub TxtClave_Change()
 CmdGuardar.Enabled = (Len(TxtClave.Text) >= 6 And TxtClave.Text = TxtConfirmar.Text)
End Sub
Private Sub TxtConfirmar_Change()
 TxtClave_Change
End Sub