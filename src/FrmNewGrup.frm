VERSION 5.00
Begin VB.Form FrmNewGrup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Generar Nuevo Grupo"
   ClientHeight    =   1320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5760
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1320
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdCancelar
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2070
      TabIndex        =   3
      Top             =   810
      Width           =   1725
   End
   Begin VB.CommandButton CmdNewGru 
      Caption         =   "Crear Grupo"
      Enabled         =   0
      Height          =   375
      Left            =   3915
      TabIndex        =   2
      Top             =   810
      Width           =   1725
   End
   Begin VB.TextBox TxtNewGru 
      Height          =   330
      Left            =   135
      TabIndex        =   0
      Top             =   360
      Width           =   5460
   End
   Begin VB.Label Label1 
      Caption         =   "Nuevo grupo:"
      Height          =   195
      Left            =   135
      TabIndex        =   1
      Top             =   90
      Width           =   1140
   End
End
Attribute VB_Name = "FrmNewGrup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public CodigoCreado As Long

Private Sub CmdNewGru_Click()
    On Error GoTo ErrInt
    If Len(Trim$(TxtNewGru.Text)) = 0 Or Not CmdNewGru.Enabled Then Exit Sub
    CmdNewGru.Enabled = False
    CodigoCreado = CrearGrupo(TxtNewGru.Text)
    Me.Hide
    Exit Sub
ErrInt:
    CmdNewGru.Enabled = (Len(Trim$(TxtNewGru.Text)) > 0)
    MsgBox Err.Description, vbExclamation, "Clever SoftPro"
End Sub

Private Sub Form_Load()
    TxtNewGru.MaxLength = 150
    CmdNewGru.Enabled = (Len(Trim$(TxtNewGru.Text)) > 0)
End Sub

Private Sub CmdCancelar_Click()
    CodigoCreado = 0: Me.Hide
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then KeyCode = 0: CmdCancelar_Click
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode = vbFormControlMenu Then Cancel = True: CmdCancelar_Click
End Sub

Private Sub TxtNewGru_Change()
    Dim Posicion As Long
    CmdNewGru.Enabled = (Len(Trim$(TxtNewGru.Text)) > 0)
    Posicion = TxtNewGru.SelStart
    If TxtNewGru.Text <> UCase$(TxtNewGru.Text) Then TxtNewGru.Text = UCase$(TxtNewGru.Text): TxtNewGru.SelStart = Posicion
End Sub
